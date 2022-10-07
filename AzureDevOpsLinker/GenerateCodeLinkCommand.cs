using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.VisualStudio.TeamFoundation.VersionControl;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.Threading;
using System.Data.SqlTypes;
using Microsoft.VisualStudio.TextManager.Interop;
using System.Windows.Forms;

namespace AzureDevOpsLinker
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class GenerateCodeLinkCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("ea4e2d7c-d7de-472e-8cc9-3120fb5dcfb6");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;

        /// <summary>
        /// Initializes a new instance of the <see cref="GenerateCodeLinkCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private GenerateCodeLinkCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static GenerateCodeLinkCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider => _package;

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in GenerateCodeLinkCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new GenerateCodeLinkCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void Execute(object sender, EventArgs e)
        {
            Flurl.Url url = default;
            Exception exception = default;

            if (await ServiceProvider.GetServiceAsync(typeof(DTE)) is DTE dte)
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                if (dte.GetObject("Microsoft.VisualStudio.TeamFoundation.VersionControl.VersionControlExt") is VersionControlExt versionControlExt
                    && versionControlExt.SolutionWorkspace is Workspace workspace)
                {
                    try
                    {
                        Uri serverUrl = workspace.VersionControlServer.TeamProjectCollection.Uri;
                        string projectName = workspace.GetTeamProjectForLocalPath(dte.ActiveDocument.FullName).Name;
                        string sourceControlPath = workspace.GetServerItemForLocalItem(dte.ActiveDocument.FullName);

                        url = new Flurl.Url(serverUrl)
                            .AppendPathSegment(projectName)
                            .AppendPathSegment("_versionControl")
                            .SetQueryParam("path", sourceControlPath);

                        if (dte.ActiveDocument.Selection is TextSelection textSelection && textSelection.CurrentLine > 1)
                        {
                            int startLine = textSelection.TopLine;
                            int endLine = textSelection.CurrentLine;

                            url
                                .SetQueryParam("lineStyle", "plain")
                                .SetQueryParam("line", startLine)
                                .SetQueryParam("lineEnd", endLine)
                                .SetQueryParam("lineStartColumn", 1)
                                .SetQueryParam("lineEndColumn", endLine > startLine ? short.MaxValue : 1);
                        }
                    }
                    catch (Exception ex)
                    {
                        exception = ex;
                    }
                }
            }

            if (!string.IsNullOrEmpty(url))
            {
                System.Diagnostics.Process.Start(url);
                Clipboard.SetText(url);
            }
            else
            {
                VsShellUtilities.ShowMessageBox(_package, "There was an error generating the URL.", "Error", OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                Clipboard.SetText(exception?.ToString() ?? string.Empty);
            }
        }
    }
}
