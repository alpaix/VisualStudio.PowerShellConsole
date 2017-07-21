// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the Apache License, Version 2.0. See License.txt in the project root for license information.

using System;
using System.Diagnostics.CodeAnalysis;
using System.Management.Automation;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace Alpaix.VisualStudio.PowerShellConsole
{
    internal static class NativeMethods
    {
        // Size of VARIANTs in 32 bit systems
        public const int VariantSize = 16;

        [DllImport("Oleaut32.dll", PreserveSig = false)]
        public static extern void VariantClear(IntPtr var);

        [DllImport("user32.dll")]
        public static extern short VkKeyScanEx(char ch, IntPtr dwhkl);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetKeyboardLayout(int dwLayout);

        [Flags]
        internal enum CreduiFlags
        {
            ALWAYS_SHOW_UI = 0x80,
            COMPLETE_USERNAME = 0x800,
            DO_NOT_PERSIST = 2,
            EXCLUDE_CERTIFICATES = 8,
            EXPECT_CONFIRMATION = 0x20000,
            GENERIC_CREDENTIALS = 0x40000,
            INCORRECT_PASSWORD = 1,
            KEEP_USERNAME = 0x100000,
            PASSWORD_ONLY_OK = 0x200,
            PERSIST = 0x1000,
            REQUEST_ADMINISTRATOR = 4,
            REQUIRE_CERTIFICATE = 0x10,
            REQUIRE_SMARTCARD = 0x100,
            SERVER_CREDENTIAL = 0x4000,
            SHOW_SAVE_CHECK_BOX = 0x40,
            USERNAME_TARGET_CREDENTIALS = 0x80000,
            VALIDATE_USERNAME = 0x400
        }

        internal enum CredUiReturnCodes
        {
            ERROR_CANCELLED = 0x4c7,
            ERROR_INSUFFICIENT_BUFFER = 0x7a,
            ERROR_INVALID_ACCOUNT_NAME = 0x523,
            ERROR_INVALID_FLAGS = 0x3ec,
            ERROR_INVALID_PARAMETER = 0x57,
            ERROR_NO_SUCH_LOGON_SESSION = 0x520,
            ERROR_NOT_FOUND = 0x490,
            NO_ERROR = 0
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct CreduiInfo
        {
            public int cbSize;
            public IntPtr hwndParent;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszMessageText;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string pszCaptionText;

            public IntPtr hbmBanner;
        }

        [DllImport("credui", EntryPoint = "CredUIPromptForCredentialsW", CharSet = CharSet.Unicode)]
        private static extern CredUiReturnCodes CredUIPromptForCredentials(ref CreduiInfo pUiInfo, string pszTargetName, IntPtr reserved, int dwAuthError, StringBuilder pszUserName, int ulUserNameMaxChars, StringBuilder pszPassword, int ulPasswordMaxChars, ref int pfSave, CreduiFlags dwFlags);

        [SuppressMessage(
            "Microsoft.Reliability",
            "CA2000:Dispose objects before losing scope",
            Justification = "Caller's responsibility to dispose.")]
        internal static PSCredential CredUIPromptForCredentials(
            string caption, string message, string userName, string targetName,
            PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options,
            IntPtr parentHwnd = default(IntPtr))
        {
            PSCredential credential = null;

            var info = new CreduiInfo
            {
                pszCaptionText = caption,
                pszMessageText = message
            };

            var pszUserName = new StringBuilder(userName, 0x201);
            var pszPassword = new StringBuilder(0x100);
            int pfSave = Convert.ToInt32(false);
            info.cbSize = Marshal.SizeOf(info);
            info.hwndParent = parentHwnd;

            var dwFlags = CreduiFlags.DO_NOT_PERSIST;
            if ((allowedCredentialTypes & PSCredentialTypes.Domain) != PSCredentialTypes.Domain)
            {
                dwFlags |= CreduiFlags.GENERIC_CREDENTIALS;
                if ((options & PSCredentialUIOptions.AlwaysPrompt) == PSCredentialUIOptions.AlwaysPrompt)
                {
                    dwFlags |= CreduiFlags.ALWAYS_SHOW_UI;
                }
            }

            var codes = CredUiReturnCodes.ERROR_INVALID_PARAMETER;

            if ((pszUserName.Length <= 0x201)
                && (pszPassword.Length <= 0x100))
            {
                codes = CredUIPromptForCredentials(
                    ref info, targetName, IntPtr.Zero, 0, pszUserName,
                    0x201, pszPassword, 0x100, ref pfSave, dwFlags);
            }

            if (codes == CredUiReturnCodes.NO_ERROR)
            {
                string providedUserName = pszUserName.ToString();
                var providedPassword = new SecureString();

                for (int i = 0; i < pszPassword.Length; i++)
                {
                    providedPassword.AppendChar(pszPassword[i]);
                    pszPassword[i] = '\0';
                }
                providedPassword.MakeReadOnly();

                if (!string.IsNullOrEmpty(providedUserName))
                {
                    credential = new PSCredential(providedUserName, providedPassword);
                }
            }
            return credential;
        }
    }
}
