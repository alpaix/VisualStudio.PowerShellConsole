using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Threading;
using Alpaix.VisualStudio.PowerShellConsole;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;

namespace Alpaix.SampleConsoleExtension
{
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
    [Guid("d3decc10-ae82-4b9b-b52e-20efb53e5287")]
    public class PowerShellConsoleToolWindow : ToolWindowPane, IOleCommandTarget
    {
        private ConsoleContainer _consoleParentPane;
        private FrameworkElement _pendingFocusPane;
        private IVsTextView _vsTextView;
        private IWpfConsole _wpfConsole;

        /// <summary>
        /// Get the WpfConsole of the active host.
        /// </summary>
        private IWpfConsole WpfConsole
        {
            get
            {
                if (_wpfConsole == null)
                {
                    var cm = GetService(typeof(SComponentModel)) as IComponentModel;
                    var pcw = cm.GetService<IPowerConsoleWindow>();
                    _wpfConsole = pcw.ActiveHostInfo.WpfConsole;
                }

                return _wpfConsole;
            }
        }

        /// <summary>
        /// Get the VsTextView of current WpfConsole if exists.
        /// </summary>
        private IVsTextView VsTextView
        {
            get
            {
                if (_vsTextView == null
                    && _wpfConsole != null)
                {
                    _vsTextView = (IVsTextView)(WpfConsole.VsTextView);
                }
                return _vsTextView;
            }
        }

        /// <summary>
        /// Get the parent pane of console panes. This serves as the Content of this tool window.
        /// </summary>
        private ConsoleContainer ConsoleParentPane
        {
            get
            {
                if (_consoleParentPane == null)
                {
                    _consoleParentPane = new ConsoleContainer();
                }
                return _consoleParentPane;
            }
        }

        private FrameworkElement PendingFocusPane
        {
            get => _pendingFocusPane;
            set
            {
                if (_pendingFocusPane != null)
                {
                    _pendingFocusPane.Loaded -= PendingFocusPane_Loaded;
                }
                _pendingFocusPane = value;
                if (_pendingFocusPane != null)
                {
                    _pendingFocusPane.Loaded += PendingFocusPane_Loaded;
                }
            }
        }

        private IVsUIShell VsUIShell => GetService(typeof(SVsUIShell)) as IVsUIShell;

        public override object Content
        {
            get { return ConsoleParentPane; }
            set { base.Content = value; }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PowerShellConsoleToolWindow"/> class.
        /// </summary>
        public PowerShellConsoleToolWindow() : base(null)
        {
            Caption = "PowerShell Console";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            //Content = new PowerShellConsoleToolWindowControl();
        }

        public override void OnToolWindowCreated()
        {
            // Register key bindings to use in the editor
            var windowFrame = (IVsWindowFrame)Frame;
            var cmdUi = VSConstants.GUID_TextEditorFactory;
            windowFrame.SetGuidProperty((int)__VSFPROPID.VSFPROPID_InheritKeyBindings, ref cmdUi);

            // pause for a tiny moment to let the tool window open before initializing the host
            var timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(10)
            };

            timer.Tick += (o, e) =>
            {
                // all exceptions from the timer thread should be caught to avoid crashing VS
                try
                {
                    LoadConsoleEditor();
                    timer.Stop();
                }
                catch (Exception x)
                {
                    ExceptionHelper.WriteErrorToActivityLog(x);
                }
            };
            timer.Start();

            base.OnToolWindowCreated();
        }

        protected override void OnClose()
        {
            base.OnClose();

            if (_wpfConsole != null)
            {
                _wpfConsole.Dispose();
            }
        }

        /// <summary>
        /// This override allows us to forward these messages to the editor instance as well
        /// </summary>
        /// <param name="m"></param>
        /// <returns></returns>
        protected override bool PreProcessMessage(ref Message m)
        {
            var vsWindowPane = VsTextView as IVsWindowPane;
            if (vsWindowPane != null)
            {
                var pMsg = new MSG[1];
                pMsg[0].hwnd = m.HWnd;
                pMsg[0].message = (uint)m.Msg;
                pMsg[0].wParam = m.WParam;
                pMsg[0].lParam = m.LParam;

                return vsWindowPane.TranslateAccelerator(pMsg) == 0;
            }

            return base.PreProcessMessage(ref m);
        }

        private void LoadConsoleEditor()
        {
            if (WpfConsole != null)
            {
                // allow the console to start writing output
                WpfConsole.StartWritingOutput();

                var consolePane = WpfConsole.Content as FrameworkElement;
                ConsoleParentPane.AddConsoleEditor(consolePane);

                // WPF doesn't handle input focus automatically in this scenario. We
                // have to set the focus manually, otherwise the editor is displayed but
                // not focused and not receiving keyboard inputs until clicked.
                if (consolePane != null)
                {
                    PendingMoveFocus(consolePane);
                }
            }
        }

        /// <summary>
        /// Set pending focus to a console pane. At the time of setting active host,
        /// the pane (UIElement) is usually not loaded yet and can't receive focus.
        /// In this case, we need to set focus in its Loaded event.
        /// </summary>
        /// <param name="consolePane"></param>
        private void PendingMoveFocus(FrameworkElement consolePane)
        {
            if (consolePane.IsLoaded
                && PresentationSource.FromDependencyObject(consolePane) != null)
            {
                PendingFocusPane = null;
                MoveFocus(consolePane);
            }
            else
            {
                PendingFocusPane = consolePane;
            }
        }

        private void PendingFocusPane_Loaded(object sender, RoutedEventArgs e)
        {
            MoveFocus(PendingFocusPane);
            PendingFocusPane = null;
        }

        private void MoveFocus(FrameworkElement consolePane)
        {
            // TAB focus into editor (consolePane.Focus() does not work due to editor layouts)
            consolePane.MoveFocus(new TraversalRequest(FocusNavigationDirection.First));

            // Try start the console session now. This needs to be after the console
            // pane getting focus to avoid incorrect initial editor layout.
            StartConsoleSession(consolePane);
        }

        private void StartConsoleSession(FrameworkElement consolePane)
        {
            if (WpfConsole != null
                && WpfConsole.Content == consolePane
                && WpfConsole.Host != null)
            {
                try
                {
                    if (WpfConsole.Dispatcher.IsStartCompleted)
                    {
                        OnDispatcherStartCompleted();
                        // if the dispatcher was started before we reach here,
                        // it means the dispatcher has been in read-only mode (due to _startedWritingOutput = false).
                        // enable key input now.
                        WpfConsole.Dispatcher.AcceptKeyInput();
                    }
                    else
                    {
                        WpfConsole.Dispatcher.StartCompleted += (sender, args) => OnDispatcherStartCompleted();
                        WpfConsole.Dispatcher.StartWaitingKey += OnDispatcherStartWaitingKey;
                        WpfConsole.Dispatcher.Start();
                    }
                }
                catch (Exception x)
                {
                    // hide the text "initialize host" when an error occurs.
                    ConsoleParentPane.NotifyInitializationCompleted();

                    WpfConsole.WriteLine(x.GetBaseException().ToString());
                    ExceptionHelper.WriteErrorToActivityLog(x);
                }
            }
            else
            {
                ConsoleParentPane.NotifyInitializationCompleted();
            }
        }

        private void OnDispatcherStartWaitingKey(object sender, EventArgs args)
        {
            WpfConsole.Dispatcher.StartWaitingKey -= OnDispatcherStartWaitingKey;
            // we want to hide the text "initialize host..." when waiting for key input
            ConsoleParentPane.NotifyInitializationCompleted();
        }

        private void OnDispatcherStartCompleted()
        {
            WpfConsole.Dispatcher.StartWaitingKey -= OnDispatcherStartWaitingKey;

            ConsoleParentPane.NotifyInitializationCompleted();

            // force the UI to update the toolbar
            VsUIShell.UpdateCommandUI(fImmediateUpdate: 0 /* false = update UI asynchronously */);

            //NuGetEventTrigger.Instance.TriggerEvent(NuGetEvent.PackageManagerConsoleLoaded);
        }
    }
}
