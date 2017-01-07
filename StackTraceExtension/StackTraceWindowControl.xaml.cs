//------------------------------------------------------------------------------
// <copyright file="StackTraceWindowControl.xaml.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace StackTraceExtension
{
    using System.Windows.Controls;

    /// <summary>
    /// Interaction logic for StackTraceWindowControl.
    /// </summary>
    public partial class StackTraceWindowControl : UserControl
    {
        public TextBlock StackBox
        {
            get { return tbStack; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="StackTraceWindowControl"/> class.
        /// </summary>
        public StackTraceWindowControl()
        {
            this.InitializeComponent();
        }
    }
}