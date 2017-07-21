// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the Apache License, Version 2.0. See License.txt in the project root for license information.

namespace Alpaix.VisualStudio.PowerShellConsole
{
    /// <summary>
    /// MEF interface to interact with the PowerConsole tool window.
    /// </summary>
    public interface IPowerConsoleWindow
    {
        HostInfo ActiveHostInfo { get; }

        /// <summary>
        /// Show the tool window
        /// </summary>
        void Show();
    }
}
