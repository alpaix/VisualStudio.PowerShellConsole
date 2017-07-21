// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the Apache License, Version 2.0. See License.txt in the project root for license information.

using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace NuGetConsole.Implementation.PowerConsole
{
    [Export(typeof(IPowerConsoleWindow))]
    [Export(typeof(IHostInitializer))]
    public class PowerConsoleWindow : IPowerConsoleWindow, IHostInitializer, IDisposable
    {
        public const string ContentType = "PackageConsole";

        private Dictionary<string, HostInfo> _hostInfos;
        private HostInfo _activeHostInfo;

        [Import(typeof(SVsServiceProvider))]
        internal IServiceProvider ServiceProvider { get; set; }

        [Import]
        internal IWpfConsoleService WpfConsoleService { get; set; }

        [ImportMany]
        internal IEnumerable<Lazy<IHostProvider, IHostMetadata>> HostProviders { get; set; }

        [SuppressMessage("Microsoft.Reliability", "CA2000:Dispose objects before losing scope",
            Justification = "_hostInfo collection is disposed.")]
        private Dictionary<string, HostInfo> HostInfos
        {
            get
            {
                if (_hostInfos == null)
                {
                    _hostInfos = new Dictionary<string, HostInfo>();
                    foreach (var p in HostProviders)
                    {
                        var info = new HostInfo(this, p);
                        _hostInfos[info.HostName] = info;
                    }
                }
                return _hostInfos;
            }
        }

        public HostInfo ActiveHostInfo
        {
            get
            {
                if (_activeHostInfo == null)
                {
                    // we only have exactly one host, the PowerShellHost. So always choose the first and only one.
                    _activeHostInfo = HostInfos.Values.FirstOrDefault();
                }
                return _activeHostInfo;
            }
        }

        [SuppressMessage("Microsoft.VisualStudio.Threading.Analyzers", "VSTHRD010", Justification = "NuGet/Home#4833 Baseline")]
        public void Show()
        {
            var vsUIShell = ServiceProvider.GetService(typeof(SVsUIShell)) as IVsUIShell;
            if (vsUIShell != null)
            {
                //var guid = typeof(PowerConsoleToolWindow).GUID;
                var guid = Guid.Parse("d3decc10-ae82-4b9b-b52e-20efb53e5287");
                IVsWindowFrame frame;

                ErrorHandler.ThrowOnFailure(
                    vsUIShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fForceCreate, ref guid, out frame));

                if (frame != null)
                {
                    ErrorHandler.ThrowOnFailure(frame.Show());
                }
            }
        }

        public void Start()
        {
            ActiveHostInfo.WpfConsole.Dispatcher.Start();
        }

        public void SetDefaultRunspace()
        {
            ActiveHostInfo.WpfConsole.Host.SetDefaultRunspace();
        }

        void IDisposable.Dispose()
        {
            if (_hostInfos != null)
            {
                foreach (var hostInfo in _hostInfos.Values.Cast<IDisposable>())
                {
                    if (hostInfo != null)
                    {
                        hostInfo.Dispose();
                    }
                }
            }
        }
    }
}
