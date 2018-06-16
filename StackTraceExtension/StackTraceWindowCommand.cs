//------------------------------------------------------------------------------
// <copyright file="StackTraceWindowCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Navigation;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.TextManager.Interop;
using System.Runtime.InteropServices;
using StackTraceExtension.Helpers;
using Parser.Enums;

namespace StackTraceExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class StackTraceWindowCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("85e5b8d2-484b-41db-9334-90b0f7fd1fbd");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="StackTraceWindowCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private StackTraceWindowCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;
            Factory.ServiceProvider = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.ShowToolWindow, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static StackTraceWindowCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new StackTraceWindowCommand(package);
        }

        /// <summary>
        /// Shows the tool window when the menu item is clicked.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event args.</param>
        private void ShowToolWindow(object sender, EventArgs e)
        {
            // Get the instance number 0 of this tool window. This window is single instance so this instance
            // is actually the only one.
            // The last flag is set to true so that if the tool window does not exists it will be created.
            var window = (StackTraceWindow)this.package.FindToolWindow(typeof(StackTraceWindow), 0, true);
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException("Cannot create tool window");
            }

            FillStack(window.Control.StackBox);
            
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
        
        private void FillStack(TextBlock stackTextBlock)
        {
            var clipboardText = Clipboard.GetText();
            if (!string.IsNullOrEmpty(clipboardText))
            {
                var parser = Factory.GetParser();
                var output = parser.ParseStack(clipboardText);

                stackTextBlock.Inlines.Clear();
                foreach (var item in output)
                {
                    var text = item.Item1;
                    var type = item.Item2;
                    switch (type)
                    {
                        case InformationType.Method: // TODO: as hyperlink
                        case InformationType.Text:
                            stackTextBlock.Inlines.Add(text);
                            break;                        
                        case InformationType.File:
                            var hyperlink = new System.Windows.Documents.Hyperlink
                            {
                                NavigateUri = new Uri("https://www.google.ru"),
                                DataContext = type // saving type of link
                            };
                            hyperlink.Inlines.Add(text);
                            hyperlink.RequestNavigate += Hyperlink_RequestNavigate;
                            stackTextBlock.Inlines.Add(hyperlink);
                            break;
                    }                 
                }                
            }
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                var link = sender as System.Windows.Documents.Hyperlink;
                var type = link.DataContext as InformationType?;

                var fileName = string.Empty;
                var lineNumber = 0;
                var columnNumber = 0;

                var linkText = ((System.Windows.Documents.Run)link.Inlines.FirstInline).Text;
                var parser = Factory.GetParser();
                switch (type)
                {
                    case InformationType.File:
                        var solution = ServiceProvider.GetService(typeof(SVsSolution)) as IVsSolution;
                        string solutionPath, solutionFile, suoFile;
                        solution.GetSolutionInfo(out solutionPath, out solutionFile, out suoFile);
                        if (string.IsNullOrEmpty(solutionPath))
                        {
                            return;
                        }
                        var tempString = solutionPath;
                        tempString = tempString.Remove(tempString.Length - 1);
                        tempString = tempString.Substring(tempString.LastIndexOf("\\") + 1, tempString.Length - tempString.LastIndexOf("\\") - 1);

                        var relationPath = linkText.Remove(0, linkText.LastIndexOf("\\" + tempString + "\\"));
                        relationPath = relationPath.Replace("\\" + tempString + "\\", string.Empty);

                        var lineNumberString = parser.GetLineNumberString(relationPath);
                        var startIndexPrefixLineNumberRelationPath = relationPath.LastIndexOf(lineNumberString);
                        relationPath = relationPath.Remove(startIndexPrefixLineNumberRelationPath, relationPath.Length - startIndexPrefixLineNumberRelationPath); // delete line number text

                        fileName = solutionPath + relationPath;
                        lineNumber = parser.GetLineNumber(linkText);
                        columnNumber = 0;

                        NavigateTo(fileName, lineNumber, columnNumber);
                        break;
                    case InformationType.Method:
                        linkText = linkText.Replace("at ", string.Empty);   // delete en prefix
                        linkText = linkText.Replace("в ", string.Empty);    // delete ru prefix
                        linkText = linkText.Remove(linkText.Length - 1);    // delete postfix

                        var projects = ProjectUtilities.GetProjectsOfCurrentSelections();
                        Factory.GetBackgroundScanner().StopIfRunning(blockUntilDone: true);

                        Factory.GetBackgroundScanner().Stopped += (source, arg) => NavigateTo(fileName, lineNumber, columnNumber);

                        Factory.GetBackgroundScanner().Start(projects, "FirstToolWindowPackage");

                        break;
                }
            }
            catch(Exception)
            {
            }
        }


        public int NavigateTo(string fileName, int lineNumber, int columnNumber = 1)
        {
            lineNumber--;
            columnNumber--;
            int hr = VSConstants.S_OK;
            var openDoc = ServiceProvider.GetService(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
            if (openDoc == null)
            {
                return VSConstants.E_UNEXPECTED;
            }

            Microsoft.VisualStudio.OLE.Interop.IServiceProvider sp = null;
            IVsUIHierarchy hierarchy = null;
            uint itemID = 0;
            IVsWindowFrame frame = null;
            Guid viewGuid = VSConstants.LOGVIEWID_TextView;

            hr = openDoc.OpenDocumentViaProject(fileName, ref viewGuid, out sp, out hierarchy, out itemID, out frame);

            hr = frame.Show();

            IntPtr viewPtr = IntPtr.Zero;
            Guid textLinesGuid = typeof(IVsTextLines).GUID;
            hr = frame.QueryViewInterface(ref textLinesGuid, out viewPtr);

            IVsTextLines textLines = Marshal.GetUniqueObjectForIUnknown(viewPtr) as IVsTextLines;

            var textMgr = ServiceProvider.GetService(typeof(SVsTextManager)) as IVsTextManager;
            if (textMgr == null)
            {
                return VSConstants.E_UNEXPECTED;
            }

            IVsTextView textView = null;
            hr = textMgr.GetActiveView(0, textLines, out textView);

            if (textView != null)
            {
                if (lineNumber >= 0)
                {
                    textView.SetCaretPos(lineNumber, Math.Max(columnNumber, 0));
                }
            }

            return VSConstants.S_OK;
        }
    }
}