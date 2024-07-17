using EnvDTE;
using EnvDTE80;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;
using Microsoft.CodeAnalysis;
using System.Linq;
using System.IO;

namespace SyncNamespace
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ListNotSynced
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 256;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("e7155a78-fe26-4a11-afb9-151abd8fdaaf");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ListNotSynced"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ListNotSynced(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ListNotSynced Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ListNotSynced's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ListNotSynced(package, commandService);
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
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            // Get DTE Object
            var dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
            UIHierarchy uih = dte.ToolWindows.SolutionExplorer;

            Array selectedItems = (Array)uih.SelectedItems;
            if (null != selectedItems)
            {
                foreach (UIHierarchyItem selItem in selectedItems)
                {
                    if (selItem.Object is Project project)
                    {
                        string filePath = project.FileName;
                        string fileShortPath = filePath.Substring(0, filePath.LastIndexOf(@"\"));
                        string projectPath = project.Name;
                        var sw = File.AppendText($"{fileShortPath}\\different_namespaces.txt");

                        foreach (var file in Directory.GetFiles(fileShortPath, "*.cs", SearchOption.AllDirectories))
                        {
                            filePath = file;
                            fileShortPath = filePath.Substring(0, filePath.LastIndexOf(@"\"));
                            string nameSpace = $"{projectPath}{fileShortPath.Substring(fileShortPath.IndexOf(projectPath) + projectPath.Length)}"
                            .Replace(@"\", ".");

                            var tree = CSharpSyntaxTree.ParseText(File.ReadAllText(file));
                            var ns = tree.GetRoot().DescendantNodes().OfType<NamespaceDeclarationSyntax>().FirstOrDefault();

                            if (ns != null && ns.Name.ToString() != nameSpace)
                                sw.WriteLine($"{ns.Name.ToString()} {file}");
                        }

                        sw.Close();
                    }
                }
            }
        }
    }
}
