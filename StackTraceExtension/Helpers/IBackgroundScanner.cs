using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;

namespace StackTraceExtension.Helpers
{
    interface IBackgroundScanner
    {
        event EventHandler Started;
        event EventHandler Stopped;

        bool IsRunning { get; }
        void Start(IEnumerable<IVsProject> projects, string searchingWord);
        void RepeatLast();
        void StopIfRunning(bool blockUntilDone);
    }
}
