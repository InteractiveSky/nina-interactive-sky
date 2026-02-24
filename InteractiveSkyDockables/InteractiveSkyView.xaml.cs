using Microsoft.Web.WebView2.Core;
using System;
using System.ComponentModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace NINA.InteractiveSky.InteractiveSkyDockables {
    public partial class InteractiveSkyView : UserControl {
        private const string HostName = "interactive-sky.local";
        private InteractiveSkyDockable _vm;

        private readonly DispatcherTimer _heartbeat;

        // NEW: preserve last important text so heartbeat doesn't overwrite it
        private string _lastNonHeartbeatDebug = "";

        public InteractiveSkyView() {
            InitializeComponent();

            Loaded += OnLoaded;
            Unloaded += OnUnloaded;
            DataContextChanged += OnDataContextChanged;

            _heartbeat = new DispatcherTimer(DispatcherPriority.Background) {
                Interval = TimeSpan.FromMilliseconds(350)
            };
            _heartbeat.Tick += (_, __) => {
                TrySendSkyState("Heartbeat");
                // IMPORTANT: do NOT steal focus every tick (it can break clicks)
                // ForceWebViewFocus("Heartbeat");
            };
        }

        private void OnLoaded(object sender, RoutedEventArgs e) {
            _ = InitAsync();
            _heartbeat.Start();
            ForceWebViewFocus("Loaded");
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) {
            _heartbeat.Stop();
            DetachVm();
        }

        private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e) {
            DetachVm();
            _vm = DataContext as InteractiveSkyDockable;
            if (_vm != null) {
                _vm.PropertyChanged += Vm_PropertyChanged;
                TrySendSkyState("DataContextChanged");
            }
        }

        private void DetachVm() {
            if (_vm != null) {
                _vm.PropertyChanged -= Vm_PropertyChanged;
                _vm = null;
            }
        }

        private void Vm_PropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (e.PropertyName == nameof(InteractiveSkyDockable.SensorWidthMm) ||
                e.PropertyName == nameof(InteractiveSkyDockable.SensorHeightMm) ||
                e.PropertyName == nameof(InteractiveSkyDockable.FocalLengthMm) ||
                e.PropertyName == nameof(InteractiveSkyDockable.MountConnected) ||
                e.PropertyName == nameof(InteractiveSkyDockable.MountRaHours) ||
                e.PropertyName == nameof(InteractiveSkyDockable.MountDecDeg) ||
                e.PropertyName == nameof(InteractiveSkyDockable.TargetRaHours) ||
                e.PropertyName == nameof(InteractiveSkyDockable.TargetDecDeg) ||
                e.PropertyName == nameof(InteractiveSkyDockable.HasPlateSolve) ||
                e.PropertyName == nameof(InteractiveSkyDockable.RotationDeg) ||
                e.PropertyName == nameof(InteractiveSkyDockable.SolveRaHours) ||
                e.PropertyName == nameof(InteractiveSkyDockable.SolveDecDeg)) {

                TrySendSkyState("VMChanged:" + e.PropertyName);
            }
        }

        private async System.Threading.Tasks.Task InitAsync() {
            try {
                DebugText.Text = "Creating WebView2 environment...";

                var userData = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "NINA", "WebView2", "InteractiveSky_" + Environment.UserName);

                Directory.CreateDirectory(userData);

                var env = await CoreWebView2Environment.CreateAsync(
                    browserExecutableFolder: null,
                    userDataFolder: userData,
                    options: null);

                await SkyWebView.EnsureCoreWebView2Async(env);

                try {
                    SkyWebView.CoreWebView2.Settings.IsZoomControlEnabled = false;
                    SkyWebView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                    SkyWebView.CoreWebView2.Settings.AreDevToolsEnabled = true; // keep while debugging
                } catch { }

                SkyWebView.CoreWebView2.WebMessageReceived += CoreWebView2_WebMessageReceived;

                var asmDir = Path.GetDirectoryName(typeof(InteractiveSkyView).Assembly.Location)
                             ?? AppDomain.CurrentDomain.BaseDirectory;

                var wwwroot = Path.Combine(asmDir, "wwwroot");
                var indexHtml = Path.Combine(wwwroot, "index.html");

                if (!File.Exists(indexHtml)) {
                    DebugText.Text = "index.html NOT FOUND ❌ " + indexHtml;
                    SkyWebView.NavigateToString("<html><body style='background:#060511;color:white;font-family:system-ui;padding:20px'>index.html not found</body></html>");
                    return;
                }

                SkyWebView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    HostName,
                    wwwroot,
                    CoreWebView2HostResourceAccessKind.Allow);

                SkyWebView.CoreWebView2.NavigationCompleted += (_, nav) => {
                    DebugText.Text = nav.IsSuccess
                        ? $"Loaded ✅ (https://{HostName}/index.html)"
                        : $"Navigation failed ❌ {nav.WebErrorStatus}";

                    if (nav.IsSuccess) {
                        ForceWebViewFocus("NavCompleted");
                        TrySendSkyState("NavCompleted");
                    }
                };

                DebugText.Text = "Navigating…";
                SkyWebView.Source = new Uri($"https://{HostName}/index.html");
            } catch (Exception ex) {
                DebugText.Text = "WebView2 init error ❌ " + ex.Message;
            }
        }

        private void CoreWebView2_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e) {
            try {
                var json = e.WebMessageAsJson;
                using var outer = JsonDocument.Parse(json);

                JsonElement root = outer.RootElement;

                // Sometimes WebView2 delivers a JSON string:  "{...}"
                if (root.ValueKind == JsonValueKind.String) {
                    var inner = root.GetString();
                    if (string.IsNullOrWhiteSpace(inner)) return;
                    using var innerDoc = JsonDocument.Parse(inner);
                    HandleMessage(innerDoc.RootElement);
                    return;
                }

                HandleMessage(root);
            } catch (Exception ex) {
                DebugText.Text = "WebMessage error ❌ " + ex.Message;
            }
        }

        private void HandleMessage(JsonElement root) {
            if (!root.TryGetProperty("type", out var typeEl)) return;
            var type = typeEl.GetString();

            if (type == "ready") {
                DebugText.Text = "Page ready ✅ sending sky state…";
                _lastNonHeartbeatDebug = DebugText.Text;
                ForceWebViewFocus("JSReady");
                TrySendSkyState("JSReady");
                return;
            }

            if (type == "setTarget") {
                var vm = _vm ?? (DataContext as InteractiveSkyDockable);
                if (vm == null) return;

                double raH = root.TryGetProperty("raHours", out var raEl) ? raEl.GetDouble() : 0.0;
                double decD = root.TryGetProperty("decDeg", out var decEl) ? decEl.GetDouble() : 0.0;

                vm.TargetRaHours = raH;
                vm.TargetDecDeg = decD;

                DebugText.Text = $"Target set ✅  RA {raH:0.000}h  Dec {decD:0.00}°";
                _lastNonHeartbeatDebug = DebugText.Text;
                TrySendSkyState("TargetSet");
                return;
            }

            if (type == "slewToTarget") {
                var vm = _vm ?? (DataContext as InteractiveSkyDockable);
                if (vm == null) return;

                DebugText.Text = "Slew requested…";
                _lastNonHeartbeatDebug = DebugText.Text;

                _ = System.Threading.Tasks.Task.Run(async () => {
                    try { await vm.SlewToTargetAsync().ConfigureAwait(false); } catch { }
                });
                return;
            }

            DebugText.Text = $"JS msg: {type}";
            _lastNonHeartbeatDebug = DebugText.Text;
        }

        private void ForceWebViewFocus(string reason) {
            try {
                if (!Dispatcher.CheckAccess()) {
                    Dispatcher.BeginInvoke(new Action(() => ForceWebViewFocus(reason)));
                    return;
                }

                Focus();
                SkyWebView?.Focus();
                Keyboard.Focus(SkyWebView);
            } catch { }
        }

        private void TrySendSkyState(string reason) {
            try {
                if (!Dispatcher.CheckAccess()) {
                    Dispatcher.BeginInvoke(new Action(() => TrySendSkyState(reason)));
                    return;
                }

                if (SkyWebView?.CoreWebView2 == null) return;

                var vm = _vm ?? (DataContext as InteractiveSkyDockable);
                if (vm == null) {
                    DebugText.Text = $"No VM yet ({reason})";
                    _lastNonHeartbeatDebug = DebugText.Text;
                    return;
                }

                var payload = new {
                    type = "skyState",

                    sensorWidth_mm = vm.SensorWidthMm,
                    sensorHeight_mm = vm.SensorHeightMm,
                    focalLength_mm = vm.FocalLengthMm,
                    rotationDeg = vm.RotationDeg,

                    hasPlateSolve = vm.HasPlateSolve,
                    solveRaHours = vm.SolveRaHours,
                    solveDecDeg = vm.SolveDecDeg,

                    mountConnected = vm.MountConnected,
                    mountRaHours = vm.MountRaHours,
                    mountDecDeg = vm.MountDecDeg,

                    targetRaHours = vm.TargetRaHours,
                    targetDecDeg = vm.TargetDecDeg
                };

                var outJson = JsonSerializer.Serialize(payload);
                SkyWebView.CoreWebView2.PostWebMessageAsJson(outJson);

                var msg =
                    $"Sent sky ✅ ({reason})  " +
                    $"Mount: {(vm.MountConnected ? "ON" : "OFF")} RA {vm.MountRaHours:0.000}h Dec {vm.MountDecDeg:0.00}°  " +
                    $"Solve: {(vm.HasPlateSolve ? "YES" : "NO")} Rot {vm.RotationDeg:0.0}°";

                if (!string.Equals(reason, "Heartbeat", StringComparison.OrdinalIgnoreCase)) {
                    _lastNonHeartbeatDebug = msg;
                    DebugText.Text = msg;
                } else {
                    if (!string.IsNullOrWhiteSpace(_lastNonHeartbeatDebug)) {
                        DebugText.Text = _lastNonHeartbeatDebug + "\n" + msg;
                    } else {
                        DebugText.Text = msg;
                    }
                }
            } catch (Exception ex) {
                DebugText.Text = "Send sky error ❌ " + ex.Message;
                _lastNonHeartbeatDebug = DebugText.Text;
            }
        }
    }
}