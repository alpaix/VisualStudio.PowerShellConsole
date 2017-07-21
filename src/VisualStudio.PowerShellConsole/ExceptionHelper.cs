// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the Apache License, Version 2.0. See License.txt in the project root for license information.

using System;
using Microsoft.VisualStudio.Shell;

namespace Alpaix.VisualStudio.PowerShellConsole
{
    public static class ExceptionHelper
    {
        public const string LogEntrySource = "PowerShell Console";

        public static void WriteErrorToActivityLog(Exception exception)
        {
            if (exception == null)
            {
                return;
            }

            ActivityLog.LogError(LogEntrySource, exception.ToString());
        }

        public static void WriteWarningToActivityLog(Exception exception)
        {
            if (exception == null)
            {
                return;
            }

            ActivityLog.LogWarning(LogEntrySource, exception.ToString());
        }
    }
}