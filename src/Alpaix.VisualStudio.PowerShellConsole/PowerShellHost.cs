// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the Apache License, Version 2.0. See License.txt in the project root for license information.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;
using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Threading;
using NuGet.Common;
using NuGet.PackageManagement.VisualStudio;
using NuGet.VisualStudio;
using Task = System.Threading.Tasks.Task;

namespace NuGetConsole.Host.PowerShell.Implementation
{
    internal abstract class PowerShellHost : IHost, IPathExpansion, IDisposable
    {
        private readonly string _name;
        private readonly IRunspaceManager _runspaceManager;
        private const string SyncModeKey = "IsSyncMode";
        private const string CancellationTokenKey = "CancellationTokenKey";
        private readonly DTE _dte;

        private IConsole _activeConsole;
        private NuGetPSHost _nugetHost;
        // indicates whether this host has been initialized.
        // null = not initilized, true = initialized successfully, false = initialized unsuccessfully
        private bool? _initialized;
        // store the current (non-truncated) project names displayed in the project name combobox
        private string[] _projectSafeNames;

        // store the current command typed so far
        private ComplexCommand _complexCommand;

        // store the current CancellationTokenSource which will be used to cancel the operation
        // in case of abort
        private CancellationTokenSource _tokenSource;

        // store the current CancellationToken. This will be set on the private data
        private CancellationToken _token;

        protected PowerShellHost(string name, IRunspaceManager runspaceManager)
        {
            _runspaceManager = runspaceManager;

            _dte = ServiceLocator.GetInstance<DTE>();

            _name = name;
            IsCommandEnabled = true;
        }

        #region Properties

        protected Pipeline ExecutingPipeline { get; set; }

        /// <summary>
        /// The host is associated with a particular console on a per-command basis.
        /// This gets set every time a command is executed on this host.
        /// </summary>
        protected IConsole ActiveConsole
        {
            get { return _activeConsole; }
            set
            {
                _activeConsole = value;
                if (_nugetHost != null)
                {
                    _nugetHost.ActiveConsole = value;
                }
            }
        }

        public bool IsCommandEnabled { get; private set; }

        protected RunspaceDispatcher Runspace { get; private set; }

        private ComplexCommand ComplexCommand
        {
            get
            {
                if (_complexCommand == null)
                {
                    _complexCommand = new ComplexCommand((allLines, lastLine) =>
                        {
                            Collection<PSParseError> errors;
                            PSParser.Tokenize(allLines, out errors);

                            // If there is a parse error token whose END is past input END, consider
                            // it a multi-line command.
                            if (errors.Count > 0)
                            {
                                if (errors.Any(e => (e.Token.Start + e.Token.Length) >= allLines.Length))
                                {
                                    return false;
                                }
                            }

                            return true;
                        });
                }
                return _complexCommand;
            }
        }

        public string Prompt
        {
            get { return ComplexCommand.IsComplete ? EvaluatePrompt() : ">> "; }
        }

        #endregion

        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private string EvaluatePrompt()
        {
            var prompt = "PM>";

            return ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                try
                {
                    // Execute the prompt function from a worker thread, so that the UI thread is not blocked waiting
                    // on it. Note that a default prompt function as defined in Profile.ps1 will simply return
                    // a string "PM>". This will always work. However, a custom "prompt" function might call
                    // Write-Host and NuGet will explicity switch to the main thread using JTF.
                    // If the main thread was blocked then, it will consistently result in a hang.
                    var output = await Task.Run(() =>
                                        Runspace.Invoke("prompt", null, outputResults: false).FirstOrDefault());
                    if (output != null)
                    {
                        var result = output.BaseObject.ToString();
                        if (!string.IsNullOrEmpty(result))
                        {
                            prompt = result;
                        }
                    }
                }
                catch (Exception ex)
                {
                    ExceptionHelper.WriteErrorToActivityLog(ex);
                }
                return prompt;
            });
        }

        /// <summary>
        /// Doing all necessary initialization works before the console accepts user inputs
        /// </summary>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        public void Initialize(IConsole console)
        {
            ThreadHelper.JoinableTaskFactory.Run(async delegate
                {
                    ActiveConsole = console;
                    if (_initialized.HasValue)
                    {
                        if (_initialized.Value
                            && console.ShowDisclaimerHeader)
                        {
                            DisplayDisclaimerAndHelpText();
                        }
                    }
                    else
                    {
                        try
                        {
                            var result = _runspaceManager.GetRunspace(console, _name);
                            Runspace = result.Item1;
                            _nugetHost = result.Item2;

                            _initialized = true;

                            if (console.ShowDisclaimerHeader)
                            {
                                DisplayDisclaimerAndHelpText();
                            }

                            UpdateWorkingDirectory();

                            // check if PMC console is actually opened, then only hook to solution load/close events.
                            if (console is IWpfConsole)
                            {
                                // Hook up solution events
                                _solutionManager.SolutionOpened += (_, __) => HandleSolutionOpened();
                                _solutionManager.SolutionClosed += (o, e) => UpdateWorkingDirectory();
                            }

                            // Set available private data on Host
                            SetPrivateDataOnHost(false);
                        }
                        catch (Exception ex)
                        {
                            // catch all exception as we don't want it to crash VS
                            _initialized = false;
                            IsCommandEnabled = false;
                            ReportError(ex);

                            ExceptionHelper.WriteErrorToActivityLog(ex);
                        }
                    }
                });
        }

        private void HandleSolutionOpened()
        {
            // Solution opened event is raised on the UI thread
            // Go off the UI thread before calling likely expensive call of ExecuteInitScriptsAsync
            // Also, it uses semaphores, do not call it from the UI thread
            Task.Run(async () =>
            {
                UpdateWorkingDirectory();
            });
        }

        private void UpdateWorkingDirectoryAndAvailableProjects()
        {
            UpdateWorkingDirectory();
        }

        private void UpdateWorkingDirectory()
        {
            ThreadHelper.JoinableTaskFactory.Run(async () =>
            {
                await TaskScheduler.Default;

                if (Runspace.RunspaceAvailability == RunspaceAvailability.Available)
                {
                    // if there is no solution open, we set the active directory to be user profile folder
                    var targetDir = _solutionManager.IsSolutionOpen ?
                        _solutionManager.SolutionDirectory :
                        Environment.GetEnvironmentVariable("USERPROFILE");

                    Runspace.ChangePSDirectory(targetDir);
                }
            });
        }

        protected abstract bool ExecuteHost(string fullCommand, string command, params object[] inputs);

        public bool Execute(IConsole console, string command, params object[] inputs)
        {
            if (console == null)
            {
                throw new ArgumentNullException(nameof(console));
            }

            if (command == null)
            {
                throw new ArgumentNullException(nameof(command));
            }

            //NuGetEventTrigger.Instance.TriggerEvent(NuGetEvent.PackageManagerConsoleCommandExecutionBegin);
            ActiveConsole = console;

            string fullCommand;
            if (ComplexCommand.AddLine(command, out fullCommand)
                && !string.IsNullOrEmpty(fullCommand))
            {
                // create a new token source with each command since CTS aren't usable once cancelled.
                _tokenSource = new CancellationTokenSource();
                _token = _tokenSource.Token;
                return ExecuteHost(fullCommand, command, inputs);
            }

            return false; // constructing multi-line command
        }

        protected void OnExecuteCommandEnd()
        {
            //NuGetEventTrigger.Instance.TriggerEvent(NuGetEvent.PackageManagerConsoleCommandExecutionEnd);

            // dispose token source related to this current command
            _tokenSource?.Dispose();
            _token = CancellationToken.None;
        }

        public void Abort()
        {
            ExecutingPipeline?.StopAsync();
            ComplexCommand.Clear();
            try
            {
                _tokenSource?.Cancel();
            }
            catch (ObjectDisposedException)
            {
                // ObjectDisposedException is expected here, since at clear console command, tokenSource
                // would have already been disposed.
            }
        }

        protected void SetPrivateDataOnHost(bool isSync)
        {
            SetPropertyValueOnHost(SyncModeKey, isSync);
            SetPropertyValueOnHost(CancellationTokenKey, _token);
        }

        private void SetPropertyValueOnHost(string propertyName, object value)
        {
            if (_nugetHost != null)
            {
                PSPropertyInfo property = _nugetHost.PrivateData.Properties[propertyName];
                if (property == null)
                {
                    property = new PSNoteProperty(propertyName, value);
                    _nugetHost.PrivateData.Properties.Add(property);
                }
                else
                {
                    property.Value = value;
                }
            }
        }

        public void SetDefaultRunspace()
        {
            Runspace.MakeDefault();
        }

        private void DisplayDisclaimerAndHelpText()
        {
            WriteLine(String.Format(CultureInfo.CurrentCulture, Resources.PowerShellHostTitle, _nugetHost.Version));
            WriteLine();

            WriteLine(Resources.Console_HelpText);
            WriteLine();
        }

        protected void ReportError(ErrorRecord record)
        {
            WriteErrorLine(Runspace.ExtractErrorFromErrorRecord(record));
        }

        protected void ReportError(Exception exception)
        {
            exception = ExceptionUtilities.Unwrap(exception);
            WriteErrorLine(exception.Message);
        }

        private void WriteErrorLine(string message)
        {
            ActiveConsole?.Write(message + Environment.NewLine, Colors.DarkRed, null);
        }

        private void WriteLine(string message = "")
        {
            ActiveConsole?.WriteLine(message);
        }

        #region ITabExpansion

        public Task<string[]> GetExpansionsAsync(string line, string lastWord, CancellationToken token)
        {
            return GetExpansionsAsyncCore(line, lastWord, token);
        }

        protected abstract Task<string[]> GetExpansionsAsyncCore(string line, string lastWord, CancellationToken token);

        protected async Task<string[]> GetExpansionsAsyncCore(string line, string lastWord, bool isSync, CancellationToken token)
        {
            // Set the _token object to the CancellationToken passed in, so that the Private Data can be set with this token
            // Powershell cmdlets will pick up the CancellationToken from the private data of the Host, and use it in their calls to NuGetPackageManager
            _token = token;
            string[] expansions;
            try
            {
                SetPrivateDataOnHost(isSync);
                expansions = await Task.Run(() =>
                    {
                        var query = from s in Runspace.Invoke(
                            @"$__pc_args=@();$input|%{$__pc_args+=$_};if(Test-Path Function:\TabExpansion2){(TabExpansion2 $__pc_args[0] $__pc_args[0].length).CompletionMatches|%{$_.CompletionText}}else{TabExpansion $__pc_args[0] $__pc_args[1]};Remove-Variable __pc_args -Scope 0;",
                            new[] { line, lastWord },
                            outputResults: false)
                                    select (s == null ? null : s.ToString());
                        return query.ToArray();
                    }, _token);
            }
            finally
            {
                // Set the _token object to the CancellationToken passed in, so that the Private Data can be set correctly
                _token = CancellationToken.None;
            }

            return expansions;
        }

        #endregion

        #region IPathExpansion

        public Task<SimpleExpansion> GetPathExpansionsAsync(string line, CancellationToken token)
        {
            return GetPathExpansionsAsyncCore(line, token);
        }

        protected abstract Task<SimpleExpansion> GetPathExpansionsAsyncCore(string line, CancellationToken token);

        protected async Task<SimpleExpansion> GetPathExpansionsAsyncCore(string line, bool isSync, CancellationToken token)
        {
            // Set the _token object to the CancellationToken passed in, so that the Private Data can be set with this token
            // Powershell cmdlets will pick up the CancellationToken from the private data of the Host, and use it in their calls to NuGetPackageManager
            _token = token;
            SetPropertyValueOnHost(CancellationTokenKey, _token);
            var simpleExpansion = await Task.Run(() =>
                {
                    PSObject expansion = Runspace.Invoke(
                        "$input|%{$__pc_args=$_}; _TabExpansionPath $__pc_args; Remove-Variable __pc_args -Scope 0",
                        new object[] { line },
                        outputResults: false).FirstOrDefault();
                    if (expansion != null)
                    {
                        int replaceStart = (int)expansion.Properties["ReplaceStart"].Value;
                        IList<string> paths = ((IEnumerable<object>)expansion.Properties["Paths"].Value).Select(o => o.ToString()).ToList();
                        return new SimpleExpansion(replaceStart, line.Length - replaceStart, paths);
                    }

                    return null;
                }, token);

            _token = CancellationToken.None;
            SetPropertyValueOnHost(CancellationTokenKey, _token);
            return simpleExpansion;
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            Runspace?.Dispose();
        }

        #endregion
    }
}
