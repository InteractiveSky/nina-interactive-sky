using NINA.Equipment.Equipment.MyCamera;
using NINA.Equipment.Interfaces.Mediator;
using NINA.Equipment.Interfaces.ViewModel;
using NINA.Profile.Interfaces;
using NINA.WPF.Base.ViewModel;
using System;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace NINA.InteractiveSky.InteractiveSkyDockables {
    [Export(typeof(IDockableVM))]
    public class InteractiveSkyDockable : DockableVM, ICameraConsumer, IDisposable, INotifyPropertyChanged {
        private readonly IProfileService _profileService;
        private readonly ICameraMediator _cameraMediator;

        // --------- IMPORTANT ----------
        // Compile-safe MEF string-contract imports.
        // NINA will inject the correct one at runtime (if present).
        // ------------------------------
        [Import("NINA.Equipment.Interfaces.Mediator.ITelescopeMediator", AllowDefault = true)]
        public object TelescopeMediator_ByName { get; set; }

        [Import("NINA.Equipment.Interfaces.Mediator.IMountMediator", AllowDefault = true)]
        public object MountMediator_ByName { get; set; }

        [Import("NINA.Equipment.Interfaces.Mediator.ITelescopeVMMediator", AllowDefault = true)]
        public object TelescopeVmMediator_ByName { get; set; }

        // --------- Plate Solve (NINA 3.x) ----------
        // Keep compile-safe; resolve at runtime by capability.
        [Import("NINA.PlateSolve.Interfaces.Mediator.IPlateSolveMediator", AllowDefault = true)]
        public object PlateSolveMediator_ByName { get; set; }

        [Import("NINA.Imaging.Interfaces.Mediator.IImageSolverMediator", AllowDefault = true)]
        public object ImageSolverMediator_ByName { get; set; }

        [Import("NINA.PlateSolve.Interfaces.Mediator.IPlateSolveVMMediator", AllowDefault = true)]
        public object PlateSolveVmMediator_ByName { get; set; }
        // -------------------------------------------

        private object _mountMediator; // resolved runtime mediator (BEST candidate)

        // Plate solve runtime provider
        private object _solveMediator;
        private bool _solveEventsHooked = false;
        private object _solveEventDelegateKeepAlive = null;
        private PropertyChangedEventHandler _solveInpcHandlerKeepAlive = null;

        // Event-hook state (keep delegates alive)
        private bool _mountEventsHooked = false;
        private object _mountEventDelegateKeepAlive = null;
        private PropertyChangedEventHandler _mountInpcHandlerKeepAlive = null;

        // ---- Public optics (mm) ----
        public double SensorWidthMm { get; private set; } = 0;
        public double SensorHeightMm { get; private set; } = 0;
        public double FocalLengthMm { get; private set; } = 800.0;

        // ---- Plate Solve state ----
        public bool HasPlateSolve { get; private set; } = false;

        // Camera angle / PA in degrees (north-up reference)
        public double RotationDeg { get; private set; } = 0.0;

        // Optional: expose last solved center (hours / degrees)
        public double SolveRaHours { get; private set; } = 0.0;
        public double SolveDecDeg { get; private set; } = 0.0;

        // ---- Mount / Target state ----
        public bool MountConnected { get; private set; } = false;
        public double MountRaHours { get; private set; } = 0.0; // hours
        public double MountDecDeg { get; private set; } = 0.0;  // degrees

        public double TargetRaHours { get; set; } = 0.0; // hours
        public double TargetDecDeg { get; set; } = 0.0;  // degrees

        // Cache file so it still works offline
        private readonly string _cachePath;

        // Background tasks control
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();

        [ImportingConstructor]
        public InteractiveSkyDockable(IProfileService profileService, ICameraMediator cameraMediator) : base(profileService) {
            _profileService = profileService;
            _cameraMediator = cameraMediator;

            Title = "Interactive Sky";

            // Optional icon
            try {
                var dict = new ResourceDictionary {
                    Source = new Uri(
                        "NINA.InteractiveSky;component/InteractiveSkyDockables/InteractiveSkyDockableTemplates.xaml",
                        UriKind.RelativeOrAbsolute)
                };

                if (dict.Contains("NINA.InteractiveSky_AltitudeSVG")) {
                    ImageGeometry = (System.Windows.Media.GeometryGroup)dict["NINA.InteractiveSky_AltitudeSVG"];
                    ImageGeometry.Freeze();
                }
            } catch { }

            // Cache location
            _cachePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "NINA", "InteractiveSky", "last_camera.json");
            Directory.CreateDirectory(Path.GetDirectoryName(_cachePath)!);

            LoadCache();
            RefreshFocalFromProfile();

            // Subscribe to camera updates
            _cameraMediator.RegisterConsumer(this);

            // Refresh focal when profile changes (event names differ by version)
            SafeSubscribeEvent(_profileService, "ActiveProfileChanged", (_, __) => RefreshFocalFromProfile());
            SafeSubscribeEvent(_profileService, "ProfileChanged", (_, __) => RefreshFocalFromProfile());
            SafeSubscribeEvent(_profileService, "LocationChanged", (_, __) => RefreshFocalFromProfile());

            // Resolve mount mediator + hook events (best effort)
            ResolveMountMediator();
            HookMountEventsIfPossible();
            RefreshMountFromMediator(); // initial attempt

            // Resolve plate solve provider + hook events + initial read
            ResolveSolveMediator();
            HookSolveEventsIfPossible();
            RefreshSolveFromMediator(); // initial attempt

            // Periodic focal refresh (rare changes; safe)
            _ = Task.Run(async () => {
                while (!_cts.IsCancellationRequested) {
                    try { RefreshFocalFromProfile(); } catch { }
                    await Task.Delay(2500, _cts.Token).ConfigureAwait(false);
                }
            }, _cts.Token);

            // Hybrid mount watchdog
            _ = Task.Run(async () => {
                while (!_cts.IsCancellationRequested) {
                    try {
                        ResolveMountMediator();

                        if (!_mountEventsHooked) {
                            HookMountEventsIfPossible();

                            if (!_mountEventsHooked) {
                                RefreshMountFromMediator();
                            }
                        }
                    } catch { }

                    await Task.Delay(500, _cts.Token).ConfigureAwait(false);
                }
            }, _cts.Token);

            // Hybrid plate-solve watchdog (event-driven best effort + fallback polling)
            _ = Task.Run(async () => {
                while (!_cts.IsCancellationRequested) {
                    try {
                        ResolveSolveMediator();

                        if (!_solveEventsHooked) {
                            HookSolveEventsIfPossible();
                        }

                        // Always attempt refresh (cheap + ensures compatibility)
                        RefreshSolveFromMediator();
                    } catch { }

                    await Task.Delay(800, _cts.Token).ConfigureAwait(false);
                }
            }, _cts.Token);
        }

        // =========================
        // IMPORTANT: resolve by capability (slew-capable mediator wins)
        // =========================
        private void ResolveMountMediator() {
            var candidates = new[] {
                TelescopeMediator_ByName,
                MountMediator_ByName,
                TelescopeVmMediator_ByName
            }.Where(x => x != null).Distinct().ToArray();

            if (candidates.Length == 0) {
                _mountMediator = null;
                return;
            }

            // Pick best candidate by scoring its public API.
            object best = null;
            int bestScore = int.MinValue;

            foreach (var c in candidates) {
                int score = ScoreMediator(c);
                if (score > bestScore) { bestScore = score; best = c; }
            }

            // Only replace if changed (avoid re-hooking churn)
            if (!ReferenceEquals(_mountMediator, best)) {
                _mountMediator = best;
                _mountEventsHooked = false; // force re-hook on new object
            }
        }

        private static int ScoreMediator(object mediator) {
            try {
                if (mediator == null) return int.MinValue;
                var t = mediator.GetType();
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance);

                int score = 0;

                bool hasAnySlew = methods.Any(m => m.Name.IndexOf("Slew", StringComparison.OrdinalIgnoreCase) >= 0);
                if (hasAnySlew) score += 5000;

                // Very strong signals
                if (methods.Any(m => m.Name.Contains("SlewToCoordinates", StringComparison.OrdinalIgnoreCase))) score += 3000;
                if (methods.Any(m => m.Name.Contains("SlewToRaDec", StringComparison.OrdinalIgnoreCase))) score += 3000;
                if (methods.Any(m => m.Name.Contains("SlewToTopocentric", StringComparison.OrdinalIgnoreCase))) score += 3000;

                // Bonus signals
                if (methods.Any(m => m.Name.Contains("Abort", StringComparison.OrdinalIgnoreCase))) score += 300;
                if (methods.Any(m => m.Name.Contains("Stop", StringComparison.OrdinalIgnoreCase))) score += 300;
                if (methods.Any(m => m.Name.Contains("Sync", StringComparison.OrdinalIgnoreCase))) score += 200;

                // Connected property bonus
                if (t.GetProperty("Connected", BindingFlags.Public | BindingFlags.Instance) != null) score += 100;
                if (t.GetProperty("IsConnected", BindingFlags.Public | BindingFlags.Instance) != null) score += 100;

                return score;
            } catch {
                return 0;
            }
        }

        // =========================
        // Plate Solve: resolve by capability (last-result provider wins)
        // =========================
        private void ResolveSolveMediator() {
            var candidates = new[] {
                PlateSolveMediator_ByName,
                ImageSolverMediator_ByName,
                PlateSolveVmMediator_ByName
            }.Where(x => x != null).Distinct().ToArray();

            if (candidates.Length == 0) {
                _solveMediator = null;
                return;
            }

            object best = null;
            int bestScore = int.MinValue;

            foreach (var c in candidates) {
                int score = ScoreSolveProvider(c);
                if (score > bestScore) { bestScore = score; best = c; }
            }

            if (!ReferenceEquals(_solveMediator, best)) {
                _solveMediator = best;
                _solveEventsHooked = false; // re-hook if provider changed
            }
        }

        private static int ScoreSolveProvider(object provider) {
            try {
                if (provider == null) return int.MinValue;
                var t = provider.GetType();
                var props = t.GetProperties(BindingFlags.Public | BindingFlags.Instance);
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance);
                var events = t.GetEvents(BindingFlags.Public | BindingFlags.Instance);

                int score = 0;

                // Strong signals: last solve result availability
                if (props.Any(p => p.Name.Equals("LastPlateSolveResult", StringComparison.OrdinalIgnoreCase))) score += 5000;
                if (props.Any(p => p.Name.Equals("LastSolveResult", StringComparison.OrdinalIgnoreCase))) score += 4500;
                if (props.Any(p => p.Name.Equals("LastResult", StringComparison.OrdinalIgnoreCase))) score += 4000;
                if (props.Any(p => p.Name.IndexOf("SolveResult", StringComparison.OrdinalIgnoreCase) >= 0)) score += 2500;

                // Events signals
                if (events.Any(e => e.Name.IndexOf("PlateSolve", StringComparison.OrdinalIgnoreCase) >= 0)) score += 1200;
                if (events.Any(e => e.Name.IndexOf("Solve", StringComparison.OrdinalIgnoreCase) >= 0)) score += 800;
                if (events.Any(e => e.Name.IndexOf("Completed", StringComparison.OrdinalIgnoreCase) >= 0)) score += 600;

                // Methods signals
                if (methods.Any(m => m.Name.IndexOf("Solve", StringComparison.OrdinalIgnoreCase) >= 0)) score += 200;

                // INPC bonus
                if (typeof(INotifyPropertyChanged).IsAssignableFrom(t)) score += 400;

                return score;
            } catch {
                return 0;
            }
        }

        // Called by NINA when camera info changes
        public void UpdateDeviceInfo(CameraInfo deviceInfo) {
            try {
                if (deviceInfo == null) return;
                if (!deviceInfo.Connected) return;

                double pxUm = ReadDouble(deviceInfo, new[] { "PixelSize", "PixelSizeUm", "PixelSizeMicron", "PixelSizeMicrons" });
                if (pxUm <= 0) {
                    double pxX = ReadDouble(deviceInfo, new[] { "PixelSizeX", "PixelSizeXUm" });
                    double pxY = ReadDouble(deviceInfo, new[] { "PixelSizeY", "PixelSizeYUm" });
                    if (pxX > 0 && pxY > 0) pxUm = (pxX + pxY) / 2.0;
                    else if (pxX > 0) pxUm = pxX;
                    else if (pxY > 0) pxUm = pxY;
                }

                double xPx = ReadDouble(deviceInfo, new[] { "XSize", "Width", "ResolutionX" });
                double yPx = ReadDouble(deviceInfo, new[] { "YSize", "Height", "ResolutionY" });

                if (pxUm > 0 && xPx > 0 && yPx > 0) {
                    var sw = (pxUm * xPx) / 1000.0;
                    var sh = (pxUm * yPx) / 1000.0;
                    SetSensor(sw, sh);
                    SaveCache();
                } else {
                    double sw = ReadDouble(deviceInfo, new[] { "SensorXSize", "SensorXSizeMm", "SensorWidth", "SensorWidthMm" });
                    double sh = ReadDouble(deviceInfo, new[] { "SensorYSize", "SensorYSizeMm", "SensorHeight", "SensorHeightMm" });
                    if (sw > 0 && sh > 0) {
                        SetSensor(sw, sh);
                        SaveCache();
                    }
                }
            } catch { }
        }

        private void SetSensor(double sw, double sh) {
            bool changed = false;

            if (sw > 0 && Math.Abs(sw - SensorWidthMm) > 0.0005) { SensorWidthMm = sw; changed = true; }
            if (sh > 0 && Math.Abs(sh - SensorHeightMm) > 0.0005) { SensorHeightMm = sh; changed = true; }

            if (changed) {
                RaiseOnUi(nameof(SensorWidthMm));
                RaiseOnUi(nameof(SensorHeightMm));
            }
        }

        public void Dispose() {
            try { _cts.Cancel(); } catch { }

            try {
                if (_mountMediator is INotifyPropertyChanged inpc && _mountInpcHandlerKeepAlive != null) {
                    inpc.PropertyChanged -= _mountInpcHandlerKeepAlive;
                }
            } catch { }

            try {
                if (_solveMediator is INotifyPropertyChanged inpc2 && _solveInpcHandlerKeepAlive != null) {
                    inpc2.PropertyChanged -= _solveInpcHandlerKeepAlive;
                }
            } catch { }

            try { _cameraMediator.RemoveConsumer(this); } catch { }
        }

        private void RefreshFocalFromProfile() {
            var p = _profileService?.ActiveProfile;
            if (p == null) return;

            var focal = TryReadFocalLengthFromProfile(p);
            if (focal > 0 && Math.Abs(focal - FocalLengthMm) > 0.001) {
                FocalLengthMm = focal;
                RaiseOnUi(nameof(FocalLengthMm));
            }
        }

        // ----------------- Plate Solve updates (event-driven best-effort + fallback) -----------------

        private void HookSolveEventsIfPossible() {
            if (_solveEventsHooked) return;
            if (_solveMediator == null) return;

            try {
                if (_solveMediator is INotifyPropertyChanged inpc) {
                    _solveInpcHandlerKeepAlive = (_, __) => {
                        try { RefreshSolveFromMediator(); } catch { }
                    };
                    inpc.PropertyChanged -= _solveInpcHandlerKeepAlive;
                    inpc.PropertyChanged += _solveInpcHandlerKeepAlive;
                    _solveEventsHooked = true;
                    return;
                }

                EventHandler handler = (_, __) => {
                    try { RefreshSolveFromMediator(); } catch { }
                };

                // Try common event names across builds
                if (TryHookEvent(_solveMediator, "PlateSolveCompleted", handler, ref _solveEventDelegateKeepAlive) ||
                    TryHookEvent(_solveMediator, "SolveCompleted", handler, ref _solveEventDelegateKeepAlive) ||
                    TryHookEvent(_solveMediator, "SolverCompleted", handler, ref _solveEventDelegateKeepAlive) ||
                    TryHookEvent(_solveMediator, "ResultChanged", handler, ref _solveEventDelegateKeepAlive) ||
                    TryHookEvent(_solveMediator, "LastResultChanged", handler, ref _solveEventDelegateKeepAlive) ||
                    TryHookEvent(_solveMediator, "PlateSolveResultChanged", handler, ref _solveEventDelegateKeepAlive) ||
                    TryHookEvent(_solveMediator, "SolutionChanged", handler, ref _solveEventDelegateKeepAlive)) {
                    _solveEventsHooked = true;
                    return;
                }
            } catch { }
        }

        private void RefreshSolveFromMediator() {
            if (_solveMediator == null) {
                SetNoSolve();
                return;
            }

            try {
                object result = null;

                // Common property patterns
                result = GetProp(_solveMediator, "LastPlateSolveResult");
                if (result == null) result = GetProp(_solveMediator, "LastSolveResult");
                if (result == null) result = GetProp(_solveMediator, "LastResult");
                if (result == null) result = GetProp(_solveMediator, "SolveResult");
                if (result == null) result = GetProp(_solveMediator, "Result");

                // Some mediators expose result via Info/State container
                if (result == null) {
                    var info = GetProp(_solveMediator, "Info") ?? GetProp(_solveMediator, "State") ?? GetProp(_solveMediator, "DeviceInfo");
                    if (info != null) {
                        result = GetProp(info, "LastPlateSolveResult") ?? GetProp(info, "LastSolveResult") ?? GetProp(info, "LastResult");
                    }
                }

                if (result == null) {
                    SetNoSolve();
                    return;
                }

                // Determine success (different names exist)
                bool success =
                    ReadBool(result, new[] { "Success", "Solved", "IsSuccess", "IsSolved", "PlateSolved" });

                // If no explicit success prop exists, we accept "has rotation + has RA/Dec" as solved.
                double rot =
                    ReadDouble(result, new[] { "PositionAngle", "Rotation", "RotationDeg", "PA", "Angle" });

                // Read RA/Dec (hours preferred)
                double raHours = ReadDouble(result, new[] { "RightAscension", "RA", "Ra", "RightAscensionHours", "RaHours" });
                if (raHours == 0) {
                    double raDeg = ReadDouble(result, new[] { "RightAscensionDegrees", "RaDegrees", "RADegrees", "RightAscensionDeg" });
                    if (raDeg != 0) raHours = raDeg / 15.0;
                }

                double decDeg = ReadDouble(result, new[] { "Declination", "DEC", "Dec", "DeclinationDegrees", "DecDegrees" });

                bool inferredSolved = (rot != 0 || (rot == 0 && ReadBool(result, new[] { "Success", "Solved", "IsSuccess", "IsSolved" })))
                                      && (raHours != 0 || decDeg != 0);

                if (!success && !inferredSolved) {
                    SetNoSolve();
                    return;
                }

                // Update fields (only raise if changed to reduce UI churn)
                bool anyChanged = false;

                if (HasPlateSolve != true) { HasPlateSolve = true; anyChanged = true; RaiseOnUi(nameof(HasPlateSolve)); }

                if (double.IsFinite(rot) && Math.Abs(rot - RotationDeg) > 0.0001) {
                    RotationDeg = rot;
                    anyChanged = true;
                    RaiseOnUi(nameof(RotationDeg));
                }

                if (double.IsFinite(raHours) && Math.Abs(raHours - SolveRaHours) > 0.000001) {
                    SolveRaHours = raHours;
                    anyChanged = true;
                    RaiseOnUi(nameof(SolveRaHours));
                }

                if (double.IsFinite(decDeg) && Math.Abs(decDeg - SolveDecDeg) > 0.000001) {
                    SolveDecDeg = decDeg;
                    anyChanged = true;
                    RaiseOnUi(nameof(SolveDecDeg));
                }

                // If solved but rotation is exactly 0, we keep it (some solves can be ~0),
                // but your UI overlay + warning uses HasPlateSolve, not rotation.
                _ = anyChanged;
            } catch {
                SetNoSolve();
            }
        }

        private void SetNoSolve() {
            if (HasPlateSolve != false) {
                HasPlateSolve = false;
                RaiseOnUi(nameof(HasPlateSolve));
            }
        }

        // ----------------- Mount updates (event-driven best-effort + fallback) -----------------

        private void HookMountEventsIfPossible() {
            if (_mountEventsHooked) return;
            if (_mountMediator == null) return;

            try {
                if (_mountMediator is INotifyPropertyChanged inpc) {
                    _mountInpcHandlerKeepAlive = (_, __) => {
                        try { RefreshMountFromMediator(); } catch { }
                    };
                    inpc.PropertyChanged -= _mountInpcHandlerKeepAlive;
                    inpc.PropertyChanged += _mountInpcHandlerKeepAlive;
                    _mountEventsHooked = true;
                    return;
                }

                EventHandler handler = (_, __) => {
                    try { RefreshMountFromMediator(); } catch { }
                };

                if (TryHookEvent(_mountMediator, "InfoChanged", handler, ref _mountEventDelegateKeepAlive) ||
                    TryHookEvent(_mountMediator, "DeviceInfoChanged", handler, ref _mountEventDelegateKeepAlive) ||
                    TryHookEvent(_mountMediator, "TelescopeInfoChanged", handler, ref _mountEventDelegateKeepAlive) ||
                    TryHookEvent(_mountMediator, "PositionChanged", handler, ref _mountEventDelegateKeepAlive) ||
                    TryHookEvent(_mountMediator, "CoordinatesChanged", handler, ref _mountEventDelegateKeepAlive) ||
                    TryHookEvent(_mountMediator, "ConnectedChanged", handler, ref _mountEventDelegateKeepAlive) ||
                    TryHookEvent(_mountMediator, "ConnectionChanged", handler, ref _mountEventDelegateKeepAlive)) {
                    _mountEventsHooked = true;
                    return;
                }
            } catch { }
        }

        private static bool TryHookEvent(object target, string eventName, EventHandler handler, ref object keepAlive) {
            try {
                if (target == null) return false;
                var t = target.GetType();
                var ei = t.GetEvent(eventName, BindingFlags.Public | BindingFlags.Instance);
                if (ei == null) return false;

                var del = Delegate.CreateDelegate(ei.EventHandlerType, handler.Target, handler.Method);
                ei.RemoveEventHandler(target, del);
                ei.AddEventHandler(target, del);

                keepAlive = del;
                return true;
            } catch {
                return false;
            }
        }

        private void RefreshMountFromMediator() {
            if (_mountMediator == null) {
                SetMountDisconnected();
                return;
            }

            try {
                object info = null;

                info = InvokeIfExists(_mountMediator, "GetInfo");
                if (info == null) info = GetProp(_mountMediator, "Info");
                if (info == null) info = GetProp(_mountMediator, "DeviceInfo");
                if (info == null) info = GetProp(_mountMediator, "TelescopeInfo");

                if (info == null) {
                    SetMountDisconnected();
                    return;
                }

                bool connected = ReadBool(info, new[] { "Connected", "IsConnected" });

                if (connected != MountConnected) {
                    MountConnected = connected;
                    RaiseOnUi(nameof(MountConnected));
                }

                if (!connected) return;

                double raHours = ReadDouble(info, new[] { "RightAscension", "RA", "Ra", "RightAscensionHours", "RaHours" });
                if (raHours == 0) {
                    double raDeg = ReadDouble(info, new[] { "RightAscensionDegrees", "RaDegrees", "RADegrees", "RightAscensionDeg" });
                    if (raDeg != 0) raHours = raDeg / 15.0;
                }

                double decDeg = ReadDouble(info, new[] { "Declination", "DEC", "Dec", "DeclinationDegrees", "DecDegrees" });

                bool changed = false;
                if (Math.Abs(raHours - MountRaHours) > 0.000001) { MountRaHours = raHours; changed = true; }
                if (Math.Abs(decDeg - MountDecDeg) > 0.000001) { MountDecDeg = decDeg; changed = true; }

                if (changed) {
                    RaiseOnUi(nameof(MountRaHours));
                    RaiseOnUi(nameof(MountDecDeg));
                }
            } catch {
                SetMountDisconnected();
            }
        }

        private void SetMountDisconnected() {
            if (MountConnected != false) {
                MountConnected = false;
                RaiseOnUi(nameof(MountConnected));
            }
        }

        // =========================
        // FIXED: Slew using Coordinates/TopocentricCoordinates (NINA TelescopeMediator expects objects)
        // =========================
        public async Task SlewToTargetAsync() {
            ResolveMountMediator(); // re-evaluate right before slew
            if (_mountMediator == null) return;

            // Refresh state right before slewing
            try { RefreshMountFromMediator(); } catch { }

            var raHours = NormalizeHours(TargetRaHours);
            var decDeg = Math.Max(-90, Math.Min(90, TargetDecDeg));
            var raDeg = raHours * 15.0;

            await RunOnUiAsync(async () => {
                // 0) FIRST: try object-based slews (this is what your log shows is available)
                if (await TryInvokeCoordinatesBasedSlewAsync(_mountMediator, raHours, raDeg, decDeg).ConfigureAwait(false)) return;

                // 1) Then try known numeric names (some mediators have them)
                if (await TryInvokeKnownSlewNamesAsync(_mountMediator, raHours, decDeg).ConfigureAwait(false)) return;

                // 2) Try known names (degrees)
                if (await TryInvokeKnownSlewNamesAsync(_mountMediator, raDeg, decDeg).ConfigureAwait(false)) return;

                // 3) Scan any public method containing "Slew" with numeric args
                await TryInvokeSlewByScanningAsync(_mountMediator, raHours, raDeg, decDeg).ConfigureAwait(false);
            }).ConfigureAwait(false);
        }

        private static async Task<bool> TryInvokeCoordinatesBasedSlewAsync(object mediator, double raHours, double raDeg, double decDeg) {
            try {
                if (mediator == null) return false;

                // Build Coordinates (prefer degrees), fallback to TopocentricCoordinates (degrees)
                object coords = BuildAstrometryObject("NINA.Astrometry.Coordinates", raHours, raDeg, decDeg);
                object topo = BuildAstrometryObject("NINA.Astrometry.TopocentricCoordinates", raHours, raDeg, decDeg);

                // TelescopeMediator in your log: SlewToCoordinatesAsync(Coordinates coords, CancellationToken token)
                if (coords != null) {
                    object taskObj =
                        InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", coords, CancellationToken.None) ??
                        InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", coords, _NoCancelTokenOrDefault());

                    if (await AwaitIfTask(taskObj).ConfigureAwait(false)) return true;
                }

                // Overload: SlewToCoordinatesAsync(TopocentricCoordinates coords, CancellationToken token)
                if (topo != null) {
                    object taskObj =
                        InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", topo, CancellationToken.None) ??
                        InvokeIfExistsLoose(mediator, "SlewToTopocentricCoordinates", topo, CancellationToken.None);

                    if (await AwaitIfTask(taskObj).ConfigureAwait(false)) return true;
                }

                // Some builds have SlewToTopocentricCoordinates(TopocentricCoordinates coords, CancellationToken token)
                if (topo != null) {
                    object taskObj =
                        InvokeIfExistsLoose(mediator, "SlewToTopocentricCoordinates", topo, CancellationToken.None);

                    if (await AwaitIfTask(taskObj).ConfigureAwait(false)) return true;
                }
            } catch { }

            return false;
        }

        private static CancellationToken _NoCancelTokenOrDefault() {
            return CancellationToken.None;
        }

        private static async Task<bool> AwaitIfTask(object taskObj) {
            try {
                if (taskObj == null) return false;

                if (taskObj is Task t) {
                    await t.ConfigureAwait(false);

                    // If it was Task<bool> (Task`1), read Result via reflection
                    var tt = taskObj.GetType();
                    if (tt.IsGenericType && tt.GetProperty("Result") != null) {
                        try {
                            var r = tt.GetProperty("Result")?.GetValue(taskObj);
                            if (r is bool b) return b; // true = slewed; false = refused
                        } catch { }
                    }

                    // Task completed, treat as success
                    return true;
                }

                // non-task return
                return true;
            } catch {
                return false;
            }
        }

        // Builds NINA.Astrometry.Coordinates or NINA.Astrometry.TopocentricCoordinates via reflection
        // Tries constructors first; if not possible, creates empty and sets properties.
        private static object BuildAstrometryObject(string fullTypeName, double raHours, double raDeg, double decDeg) {
            try {
                var t = FindTypeAcrossLoadedAssemblies(fullTypeName);
                if (t == null) return null;

                // Try constructors (best-effort)
                var ctors = t.GetConstructors(BindingFlags.Public | BindingFlags.Instance);
                foreach (var c in ctors) {
                    var ps = c.GetParameters();

                    // Common patterns: (double ra, double dec) OR (double ra, double dec, ...)
                    if (ps.Length >= 2 && IsNumeric(ps[0].ParameterType) && IsNumeric(ps[1].ParameterType)) {
                        // Attempt degrees first, then hours
                        object[] args = new object[ps.Length];
                        args[0] = ConvertNumeric(raDeg, ps[0].ParameterType);
                        args[1] = ConvertNumeric(decDeg, ps[1].ParameterType);

                        for (int i = 2; i < ps.Length; i++) {
                            var pt = ps[i].ParameterType;
                            if (pt == typeof(bool)) { args[i] = false; continue; }
                            if (pt == typeof(CancellationToken)) { args[i] = CancellationToken.None; continue; }
                            if (pt.IsValueType) { args[i] = Activator.CreateInstance(pt); continue; }
                            args[i] = null;
                        }

                        try {
                            var obj = c.Invoke(args);
                            if (obj != null) return obj;
                        } catch { }

                        // Try hours
                        args[0] = ConvertNumeric(raHours, ps[0].ParameterType);
                        try {
                            var obj2 = c.Invoke(args);
                            if (obj2 != null) return obj2;
                        } catch { }
                    }
                }

                // If no constructor worked, create default and set common properties
                object inst = null;
                try { inst = Activator.CreateInstance(t); } catch { }
                if (inst == null) return null;

                // Try set RA/Dec properties (both deg and hours, whichever exists)
                TrySetNumericProperty(inst, new[] { "RightAscension", "RA", "Ra", "RightAscensionDegrees", "RaDegrees", "RADegrees", "RightAscensionDeg" }, raDeg);
                TrySetNumericProperty(inst, new[] { "RightAscensionHours", "RaHours" }, raHours);

                TrySetNumericProperty(inst, new[] { "Declination", "DEC", "Dec", "DeclinationDegrees", "DecDegrees" }, decDeg);

                return inst;
            } catch {
                return null;
            }
        }

        private static Type FindTypeAcrossLoadedAssemblies(string fullName) {
            try {
                // First try Type.GetType directly
                var direct = Type.GetType(fullName);
                if (direct != null) return direct;

                // Search loaded assemblies
                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies()) {
                    try {
                        var t = asm.GetType(fullName, throwOnError: false, ignoreCase: false);
                        if (t != null) return t;
                    } catch { }
                }
            } catch { }
            return null;
        }

        private static void TrySetNumericProperty(object obj, string[] propNames, double value) {
            try {
                var t = obj.GetType();
                foreach (var n in propNames) {
                    var pi = t.GetProperty(n, BindingFlags.Public | BindingFlags.Instance);
                    if (pi == null) continue;
                    if (!pi.CanWrite) continue;
                    try {
                        var converted = ConvertNumeric(value, pi.PropertyType);
                        pi.SetValue(obj, converted);
                        return;
                    } catch { }
                }
            } catch { }
        }

        private static async Task RunOnUiAsync(Func<Task> work) {
            try {
                var disp = Application.Current?.Dispatcher;
                if (disp == null) { await work().ConfigureAwait(false); return; }

                if (disp.CheckAccess()) {
                    await work().ConfigureAwait(false);
                    return;
                }

                await disp.InvokeAsync(async () => {
                    try { await work().ConfigureAwait(false); } catch { }
                });
            } catch {
                try { await work().ConfigureAwait(false); } catch { }
            }
        }

        private static async Task<bool> TryInvokeKnownSlewNamesAsync(object mediator, double ra, double dec) {
            // async variants
            object taskObj =
                InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", ra, dec) ??
                InvokeIfExistsLoose(mediator, "SlewToRaDecAsync", ra, dec) ??
                InvokeIfExistsLoose(mediator, "SlewToCoordinatesJNowAsync", ra, dec) ??
                InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", ra, dec, false) ??
                InvokeIfExistsLoose(mediator, "SlewToRaDecAsync", ra, dec, false) ??
                InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", ra, dec, true) ??
                InvokeIfExistsLoose(mediator, "SlewToRaDecAsync", ra, dec, true);

            if (taskObj is Task t) { await t.ConfigureAwait(false); return true; }

            // sync variants (some return void)
            if (TryInvokeVoidOrAny(mediator, "SlewToCoordinates", ra, dec)) return true;
            if (TryInvokeVoidOrAny(mediator, "SlewToRaDec", ra, dec)) return true;
            if (TryInvokeVoidOrAny(mediator, "SlewToCoordinatesJNow", ra, dec)) return true;

            return false;
        }

        private static async Task<bool> TryInvokeSlewByScanningAsync(object mediator, double raHours, double raDeg, double decDeg) {
            try {
                var t = mediator.GetType();
                var methods = t.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                    .Where(m => m.Name.IndexOf("Slew", StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToArray();

                foreach (var m in methods) {
                    var ps = m.GetParameters();
                    if (ps.Length < 2) continue;

                    if (!IsNumeric(ps[0].ParameterType) || !IsNumeric(ps[1].ParameterType)) continue;

                    if (await TryInvokeOneCandidateAsync(mediator, m, raHours, decDeg).ConfigureAwait(false)) return true;
                    if (await TryInvokeOneCandidateAsync(mediator, m, raDeg, decDeg).ConfigureAwait(false)) return true;
                }
            } catch { }
            return false;
        }

        private static async Task<bool> TryInvokeOneCandidateAsync(object instance, MethodInfo mi, double ra, double dec) {
            try {
                var ps = mi.GetParameters();
                var args = BuildArgs(ps, ra, dec);
                if (args == null) return false;

                object result = null;
                try { result = mi.Invoke(instance, args); } catch { return false; }

                if (result is Task t) {
                    await t.ConfigureAwait(false);
                    return true;
                }

                if (mi.ReturnType == typeof(void)) return true;
                if (result != null) return true;
            } catch { }
            return false;
        }

        private static object[] BuildArgs(ParameterInfo[] ps, double ra, double dec) {
            try {
                object[] args = new object[ps.Length];
                args[0] = ConvertNumeric(ra, ps[0].ParameterType);
                args[1] = ConvertNumeric(dec, ps[1].ParameterType);

                for (int i = 2; i < ps.Length; i++) {
                    var pt = ps[i].ParameterType;
                    if (pt == typeof(bool)) { args[i] = false; continue; }
                    if (pt == typeof(CancellationToken)) { args[i] = CancellationToken.None; continue; }
                    return null; // unknown extra param
                }

                return args;
            } catch {
                return null;
            }
        }

        private static bool IsNumeric(Type t) {
            t = Nullable.GetUnderlyingType(t) ?? t;
            return t == typeof(double) || t == typeof(float) || t == typeof(int) || t == typeof(long) || t == typeof(decimal);
        }

        private static object ConvertNumeric(double v, Type targetType) {
            targetType = Nullable.GetUnderlyingType(targetType) ?? targetType;
            if (targetType == typeof(double)) return v;
            if (targetType == typeof(float)) return (float)v;
            if (targetType == typeof(decimal)) return (decimal)v;
            if (targetType == typeof(int)) return (int)Math.Round(v);
            if (targetType == typeof(long)) return (long)Math.Round(v);
            return v;
        }

        private static double NormalizeHours(double raHours) {
            if (double.IsNaN(raHours) || double.IsInfinity(raHours)) return 0;
            while (raHours < 0) raHours += 24;
            while (raHours >= 24) raHours -= 24;
            return raHours;
        }

        // ---- Cache ----
        private void LoadCache() {
            try {
                if (!File.Exists(_cachePath)) return;
                var json = File.ReadAllText(_cachePath);
                var data = JsonSerializer.Deserialize<Cache>(json);
                if (data?.SensorWidthMm > 0 && data?.SensorHeightMm > 0) {
                    SensorWidthMm = data.SensorWidthMm;
                    SensorHeightMm = data.SensorHeightMm;
                }
            } catch { }
        }

        private void SaveCache() {
            try {
                if (SensorWidthMm <= 0 || SensorHeightMm <= 0) return;
                var json = JsonSerializer.Serialize(new Cache { SensorWidthMm = SensorWidthMm, SensorHeightMm = SensorHeightMm });
                File.WriteAllText(_cachePath, json);
            } catch { }
        }

        private sealed class Cache {
            public double SensorWidthMm { get; set; }
            public double SensorHeightMm { get; set; }
        }

        // ---- UI-safe property changed ----
        private void RaiseOnUi(string propertyName) {
            try {
                var disp = Application.Current?.Dispatcher;
                if (disp != null && !disp.CheckAccess()) {
                    disp.BeginInvoke(new Action(() => RaisePropertyChanged(propertyName)));
                } else {
                    RaisePropertyChanged(propertyName);
                }
            } catch {
                try { RaisePropertyChanged(propertyName); } catch { }
            }
        }

        // ---- Reflection helpers ----
        private static double ReadDouble(object root, string[] names) {
            if (root == null) return 0;
            var t = root.GetType();
            foreach (var name in names) {
                var pi = t.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) continue;
                try {
                    var v = pi.GetValue(root);
                    if (v == null) continue;
                    if (v is double d) return d;
                    if (v is float f) return f;
                    if (v is int i) return i;
                    if (v is long l) return l;
                    if (double.TryParse(v.ToString(), out var parsed)) return parsed;
                } catch { }
            }
            return 0;
        }

        private static bool ReadBool(object root, string[] names) {
            if (root == null) return false;
            var t = root.GetType();
            foreach (var name in names) {
                var pi = t.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) continue;
                try {
                    var v = pi.GetValue(root);
                    if (v == null) continue;
                    if (v is bool b) return b;
                    if (bool.TryParse(v.ToString(), out var parsed)) return parsed;
                } catch { }
            }
            return false;
        }

        private static object GetProp(object root, string propName) {
            if (root == null) return null;
            try {
                var pi = root.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) return null;
                return pi.GetValue(root);
            } catch {
                return null;
            }
        }

        private static object InvokeIfExists(object root, string methodName, params object[] args) {
            if (root == null) return null;
            try {
                var t = root.GetType();
                foreach (var mi in t.GetMethods(BindingFlags.Public | BindingFlags.Instance)) {
                    if (!string.Equals(mi.Name, methodName, StringComparison.Ordinal)) continue;
                    var p = mi.GetParameters();
                    if (p.Length != (args?.Length ?? 0)) continue;
                    return mi.Invoke(root, args);
                }
            } catch { }
            return null;
        }

        // Loose invoke: allow numeric coercion; same arg count
        private static object InvokeIfExistsLoose(object root, string methodName, params object[] args) {
            if (root == null) return null;
            try {
                var t = root.GetType();
                foreach (var mi in t.GetMethods(BindingFlags.Public | BindingFlags.Instance)) {
                    if (!string.Equals(mi.Name, methodName, StringComparison.Ordinal)) continue;

                    var p = mi.GetParameters();
                    if (p.Length != (args?.Length ?? 0)) continue;

                    bool ok = true;
                    for (int i = 0; i < p.Length; i++) {
                        var pt = p[i].ParameterType;
                        if (args[i] == null) continue;

                        if (IsNumeric(pt) && args[i] is IConvertible) continue;
                        if (!pt.IsInstanceOfType(args[i])) { ok = false; break; }
                    }
                    if (!ok) continue;

                    return mi.Invoke(root, args);
                }
            } catch { }
            return null;
        }

        // Treat void methods as success too (invoke throws if it fails)
        private static bool TryInvokeVoidOrAny(object root, string methodName, params object[] args) {
            if (root == null) return false;
            try {
                var t = root.GetType();
                foreach (var mi in t.GetMethods(BindingFlags.Public | BindingFlags.Instance)) {
                    if (!string.Equals(mi.Name, methodName, StringComparison.Ordinal)) continue;
                    var p = mi.GetParameters();
                    if (p.Length != (args?.Length ?? 0)) continue;
                    mi.Invoke(root, args);
                    return true;
                }
            } catch { }
            return false;
        }

        private static double TryReadFocalLengthFromProfile(object profile) {
            object tel =
                FindObject(profile, new[] { "TelescopeSettings", "AstrometrySettings", "Telescope", "OpticsSettings" }) ??
                profile;

            double f = FindDouble(tel, new[] { "FocalLength", "FocalLengthMm", "EffectiveFocalLength", "FocalLengthValue" });
            if (f > 0) return f;

            return 0;
        }

        private static object FindObject(object root, string[] names) {
            if (root == null) return null;
            var t = root.GetType();
            foreach (var name in names) {
                var pi = t.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) continue;
                try { return pi.GetValue(root); } catch { }
            }
            return null;
        }

        private static double FindDouble(object root, string[] names) {
            if (root == null) return 0;
            var t = root.GetType();
            foreach (var name in names) {
                var pi = t.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
                if (pi == null) continue;
                try {
                    var v = pi.GetValue(root);
                    if (v == null) continue;
                    if (v is double d) return d;
                    if (v is float f) return f;
                    if (v is int i) return i;
                    if (v is long l) return l;
                    if (double.TryParse(v.ToString(), out var parsed)) return parsed;
                } catch { }
            }
            return 0;
        }

        private static void SafeSubscribeEvent(object target, string eventName, EventHandler handler) {
            if (target == null) return;
            try {
                var ei = target.GetType().GetEvent(eventName, BindingFlags.Public | BindingFlags.Instance);
                if (ei == null) return;
                var del = Delegate.CreateDelegate(ei.EventHandlerType, handler.Target, handler.Method);
                ei.RemoveEventHandler(target, del);
                ei.AddEventHandler(target, del);
            } catch { }
        }
    }
}