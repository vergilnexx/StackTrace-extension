//------------------------------------------------------------------------------
// <copyright file="StackTraceWindow.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace StackTraceExtension
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.VisualStudio.Shell;
    using System.Windows.Forms;

    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    /// <remarks>
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    /// <para>
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </para>
    /// </remarks>
    [Guid("92d7323c-d008-4d1f-b5fa-de5d89c6d80a")]
    public class StackTraceWindow : ToolWindowPane
    {
        public StackTraceWindowControl Control { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="StackTraceWindow"/> class.
        /// </summary>
        public StackTraceWindow() : base(null)
        {
            this.Caption = "StackTrace Window";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            Control = new StackTraceWindowControl();
            this.Content = Control;            
        }
    }
}
