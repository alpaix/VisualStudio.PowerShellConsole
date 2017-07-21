// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the Apache License, Version 2.0. See License.txt in the project root for license information.

using System.Windows.Media;
using NuGet.VisualStudio;

namespace NuGetConsole.Host
{
    /// <summary>
    /// This host is used when PowerShell 2.0 runtime is not installed in the system. It's basically a no-op host.
    /// </summary>
    internal class UnsupportedHost : IHost
    {
        public bool IsCommandEnabled => false;

        public void Initialize(IConsole console)
        {
            // display the error message at the beginning
            console.Write(PowerShell.Resources.Host_PSNotInstalled, Colors.Red, null);
        }

        public string Prompt => string.Empty;

        public bool Execute(IConsole console, string command, object[] inputs)
        {
            return false;
        }

        public void Abort()
        {
        }

        public void SetDefaultRunspace()
        {
        }
    }
}
