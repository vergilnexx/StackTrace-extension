using Microsoft.VisualStudio.Shell.Interop;
using Parser;
using System;
using System.Collections.Generic;

namespace StackTraceExtension.Helpers
{
    internal static class Factory
    {
        private static IServiceProvider _serviceProvider;

        private static IBackgroundScanner _backgroundScanner;

        private static IStackParser _parser;

        public static IServiceProvider ServiceProvider
        {
            set
            {
                _serviceProvider = value;
                ProjectUtilities.SetServiceProvider(_serviceProvider);
            }

            get { return _serviceProvider; }
        }

        public static IBackgroundScanner GetBackgroundScanner()
        {
            if (_backgroundScanner == null)
            {
                _backgroundScanner = new BackgroundScanner(_serviceProvider);
            }
            return _backgroundScanner;
        }

        public static IStackParser GetParser()
        {
            if (_parser == null)
            {
                _parser = new StackParser();
            }
            return _parser;
        }
    }
}
