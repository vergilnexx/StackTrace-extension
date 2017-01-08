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

            GetFormatStack(window.Control.StackBox);
            
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        private void GetFormatStack(TextBlock stackTextBlock)
        {
            var clipboardText = Clipboard.GetText();
            if (!string.IsNullOrEmpty(clipboardText))
            {
                stackTextBlock.Inlines.Clear();
                stackTextBlock.Inlines.Add(clipboardText);
                var hyperlink = new System.Windows.Documents.Hyperlink {
                    NavigateUri = new Uri("https://www.google.ru")
                };
                hyperlink.Inlines.Add(clipboardText);
                hyperlink.RequestNavigate += Hyperlink_RequestNavigate;
                stackTextBlock.Inlines.Add(hyperlink);
            }
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            // TODO: calculate parameters
            var fileName = @"C:\Users\Sergey\Documents\Visual Studio 2015\Projects\FirstToolWin\FirstToolWin\FirstToolWindowPackage.cs"; 
            var lineNumber = 44;
            var columnNumber = 25;

            NavigateTo(fileName, lineNumber, columnNumber);
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
