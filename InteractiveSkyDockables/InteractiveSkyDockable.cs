// =========================
// File: InteractiveSkyDockable.cs
// =========================
using NINA.Equipment.Equipment.MyCamera;
using NINA.Equipment.Interfaces.Mediator;
using NINA.Equipment.Interfaces.ViewModel;
using NINA.Profile.Interfaces;
using NINA.WPF.Base.ViewModel;
using System;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
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
        // Different builds expose different mediators/names.
        [Import("NINA.PlateSolve.Interfaces.Mediator.IPlateSolveMediator", AllowDefault = true)]
        public object PlateSolveMediator_ByName { get; set; }

        [Import("NINA.Imaging.Interfaces.Mediator.IImageSolverMediator", AllowDefault = true)]
        public object ImageSolverMediator_ByName { get; set; }

        [Import("NINA.PlateSolve.Interfaces.Mediator.IPlateSolveVMMediator", AllowDefault = true)]
        public object PlateSolveVmMediator_ByName { get; set; }

        // EXTRA candidates (some builds rename/move these)
        [Import("NINA.Imaging.Interfaces.Mediator.IImageSolverVMMediator", AllowDefault = true)]
        public object ImageSolverVmMediator_ByName { get; set; }

        [Import("NINA.PlateSolve.Interfaces.Mediator.IImageSolverMediator", AllowDefault = true)]
        public object PlateSolveImageSolverMediator_ByName { get; set; }

        [Import("NINA.PlateSolve.Interfaces.Mediator.IImageSolverVMMediator", AllowDefault = true)]
        public object PlateSolveImageSolverVmMediator_ByName { get; set; }
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

        // --------- NEW: Log tail fallback for Plate Solve ----------
        // If mediators don't expose last solve result, we parse NINA logs:
        // "Platesolve successful: Coordinates: RA: 05:34:12; Dec: -05° 28' 12\"; ... - Position Angle: 23.50"
        private readonly NinaLogTailer _logTailer;
        // ----------------------------------------------------------

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

            // NEW: start log tailer fallback (works even if mediators don't expose results)
            _logTailer = new NinaLogTailer(_cts.Token, ApplySolveFromLog);
            _logTailer.Start();

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

                    await Task.Delay(700, _cts.Token).ConfigureAwait(false);
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

            object best = null;
            int bestScore = int.MinValue;

            foreach (var c in candidates) {
                int score = ScoreMediator(c);
                if (score > bestScore) { bestScore = score; best = c; }
            }

            if (!ReferenceEquals(_mountMediator, best)) {
                _mountMediator = best;
                _mountEventsHooked = false;
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

                if (methods.Any(m => m.Name.Contains("SlewToCoordinates", StringComparison.OrdinalIgnoreCase))) score += 3000;
                if (methods.Any(m => m.Name.Contains("SlewToRaDec", StringComparison.OrdinalIgnoreCase))) score += 3000;
                if (methods.Any(m => m.Name.Contains("SlewToTopocentric", StringComparison.OrdinalIgnoreCase))) score += 3000;

                if (methods.Any(m => m.Name.Contains("Abort", StringComparison.OrdinalIgnoreCase))) score += 300;
                if (methods.Any(m => m.Name.Contains("Stop", StringComparison.OrdinalIgnoreCase))) score += 300;
                if (methods.Any(m => m.Name.Contains("Sync", StringComparison.OrdinalIgnoreCase))) score += 200;

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
                PlateSolveVmMediator_ByName,
                ImageSolverVmMediator_ByName,
                PlateSolveImageSolverMediator_ByName,
                PlateSolveImageSolverVmMediator_ByName
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
                _solveEventsHooked = false;
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

                if (props.Any(p => p.Name.Equals("LastPlateSolveResult", StringComparison.OrdinalIgnoreCase))) score += 6000;
                if (props.Any(p => p.Name.Equals("LastSolveResult", StringComparison.OrdinalIgnoreCase))) score += 5500;
                if (props.Any(p => p.Name.Equals("LastResult", StringComparison.OrdinalIgnoreCase))) score += 5000;
                if (props.Any(p => p.Name.IndexOf("SolveResult", StringComparison.OrdinalIgnoreCase) >= 0)) score += 2500;

                if (events.Any(e => e.Name.IndexOf("PlateSolve", StringComparison.OrdinalIgnoreCase) >= 0)) score += 1200;
                if (events.Any(e => e.Name.IndexOf("Solve", StringComparison.OrdinalIgnoreCase) >= 0)) score += 800;
                if (events.Any(e => e.Name.IndexOf("Completed", StringComparison.OrdinalIgnoreCase) >= 0)) score += 600;
                if (events.Any(e => e.Name.IndexOf("Result", StringComparison.OrdinalIgnoreCase) >= 0)) score += 500;

                if (methods.Any(m => m.Name.IndexOf("Solve", StringComparison.OrdinalIgnoreCase) >= 0)) score += 200;

                if (typeof(INotifyPropertyChanged).IsAssignableFrom(t)) score += 700;

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

            try { _logTailer?.Dispose(); } catch { }
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
                    TryHookEvent(_solveMediator, "SolutionChanged", handler, ref _solveEventDelegateKeepAlive) ||
                    TryHookEvent(_solveMediator, "ImageSolved", handler, ref _solveEventDelegateKeepAlive) ||
                    TryHookEvent(_solveMediator, "SolveResultChanged", handler, ref _solveEventDelegateKeepAlive)) {
                    _solveEventsHooked = true;
                    return;
                }
            } catch { }
        }

        private void RefreshSolveFromMediator() {
            if (_solveMediator == null) {
                // do NOT force-no-solve; log tailer may still provide it
                return;
            }

            try {
                object result = FindAnySolveResultObject(_solveMediator);

                if (result == null) {
                    // do NOT force-no-solve; log tailer may still provide it
                    return;
                }

                bool success =
                    ReadBool(result, new[] { "Success", "Solved", "IsSuccess", "IsSolved", "PlateSolved", "HasSolution" });

                double rot =
                    ReadDouble(result, new[] { "PositionAngle", "PositionAngleDeg", "Rotation", "RotationDeg", "PA", "Angle" });

                double raHours = ReadRaHours(result);
                double decDeg = ReadDecDeg(result);

                if ((raHours == 0 && decDeg == 0) || double.IsNaN(raHours) || double.IsNaN(decDeg)) {
                    var coords = GetProp(result, "Coordinates") ?? GetProp(result, "Center") ?? GetProp(result, "Coo") ?? GetProp(result, "SolvedCoordinates");
                    if (coords != null) {
                        if (raHours == 0) raHours = ReadRaHours(coords);
                        if (decDeg == 0) decDeg = ReadDecDeg(coords);
                        if (rot == 0) {
                            rot = ReadDouble(coords, new[] { "PositionAngle", "Rotation", "RotationDeg", "PA", "Angle" });
                        }
                    }
                }

                rot = NormalizeAngleToDegrees(rot, result);

                bool inferredSolved = (raHours != 0 || decDeg != 0) && (success || HasAnySolveFlag(result));
                if (!success && !inferredSolved) return;

                ApplySolve(raHours, decDeg, rot);
            } catch {
                // ignore; log tailer might still work
            }
        }

        private void ApplySolve(double raHours, double decDeg, double rotDeg) {
            // If nothing is valid, ignore
            if (!double.IsFinite(raHours) || !double.IsFinite(decDeg)) return;

            // Normalize / clamp
            raHours = NormalizeHours(raHours);
            decDeg = Math.Max(-90, Math.Min(90, decDeg));

            bool anyChange = false;

            if (HasPlateSolve != true) { HasPlateSolve = true; anyChange = true; RaiseOnUi(nameof(HasPlateSolve)); }

            if (double.IsFinite(rotDeg)) {
                // normalize to 0..360
                rotDeg = rotDeg % 360.0;
                if (rotDeg < 0) rotDeg += 360.0;
                if (Math.Abs(rotDeg - RotationDeg) > 0.0001) { RotationDeg = rotDeg; anyChange = true; RaiseOnUi(nameof(RotationDeg)); }
            }

            if (Math.Abs(raHours - SolveRaHours) > 0.000001) { SolveRaHours = raHours; anyChange = true; RaiseOnUi(nameof(SolveRaHours)); }
            if (Math.Abs(decDeg - SolveDecDeg) > 0.000001) { SolveDecDeg = decDeg; anyChange = true; RaiseOnUi(nameof(SolveDecDeg)); }

            if (!anyChange) return;
        }

        private void ApplySolveFromLog(NinaSolveLine solve) {
            try {
                ApplySolve(solve.RaHours, solve.DecDeg, solve.PositionAngleDeg);
            } catch { }
        }

        private static object FindAnySolveResultObject(object mediator) {
            if (mediator == null) return null;

            object result =
                GetProp(mediator, "LastPlateSolveResult") ??
                GetProp(mediator, "LastSolveResult") ??
                GetProp(mediator, "LastResult") ??
                GetProp(mediator, "SolveResult") ??
                GetProp(mediator, "Result") ??
                GetProp(mediator, "CurrentResult");

            if (result != null) return result;

            var info =
                GetProp(mediator, "Info") ??
                GetProp(mediator, "State") ??
                GetProp(mediator, "DeviceInfo") ??
                GetProp(mediator, "LastInfo") ??
                GetProp(mediator, "SolverState") ??
                GetProp(mediator, "PlateSolveState");

            if (info != null) {
                result =
                    GetProp(info, "LastPlateSolveResult") ??
                    GetProp(info, "LastSolveResult") ??
                    GetProp(info, "LastResult") ??
                    GetProp(info, "SolveResult") ??
                    GetProp(info, "Result") ??
                    GetProp(info, "CurrentResult");
                if (result != null) return result;
            }

            var last = GetProp(mediator, "Last") ?? GetProp(mediator, "Latest");
            if (last != null) {
                result =
                    GetProp(last, "PlateSolveResult") ??
                    GetProp(last, "SolveResult") ??
                    GetProp(last, "Result");
                if (result != null) return result;
            }

            return null;
        }

        private static bool HasAnySolveFlag(object result) {
            if (result == null) return false;
            if (ReadBool(result, new[] { "Success", "Solved", "IsSuccess", "IsSolved", "PlateSolved", "HasSolution" })) return true;

            var raAny = GetAnyProp(result, new[] { "RightAscension", "RA", "Ra", "RightAscensionHours", "RaHours", "RightAscensionDegrees", "RaDegrees" });
            var decAny = GetAnyProp(result, new[] { "Declination", "DEC", "Dec", "DeclinationDegrees", "DecDegrees" });

            return raAny != null || decAny != null;
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

                double raHours = ReadRaHours(info);
                double decDeg = ReadDecDeg(info);

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
            ResolveMountMediator();
            if (_mountMediator == null) return;

            try { RefreshMountFromMediator(); } catch { }

            var raHours = NormalizeHours(TargetRaHours);
            var decDeg = Math.Max(-90, Math.Min(90, TargetDecDeg));
            var raDeg = raHours * 15.0;

            await RunOnUiAsync(async () => {
                if (await TryInvokeCoordinatesBasedSlewAsync(_mountMediator, raHours, raDeg, decDeg).ConfigureAwait(false)) return;
                if (await TryInvokeKnownSlewNamesAsync(_mountMediator, raHours, decDeg).ConfigureAwait(false)) return;
                if (await TryInvokeKnownSlewNamesAsync(_mountMediator, raDeg, decDeg).ConfigureAwait(false)) return;
                await TryInvokeSlewByScanningAsync(_mountMediator, raHours, raDeg, decDeg).ConfigureAwait(false);
            }).ConfigureAwait(false);
        }

        private static async Task<bool> TryInvokeCoordinatesBasedSlewAsync(object mediator, double raHours, double raDeg, double decDeg) {
            try {
                if (mediator == null) return false;

                object coords = BuildAstrometryObject("NINA.Astrometry.Coordinates", raHours, raDeg, decDeg);
                object topo = BuildAstrometryObject("NINA.Astrometry.TopocentricCoordinates", raHours, raDeg, decDeg);

                if (coords != null) {
                    object taskObj =
                        InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", coords, CancellationToken.None) ??
                        InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", coords, _NoCancelTokenOrDefault());

                    if (await AwaitIfTask(taskObj).ConfigureAwait(false)) return true;
                }

                if (topo != null) {
                    object taskObj =
                        InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", topo, CancellationToken.None) ??
                        InvokeIfExistsLoose(mediator, "SlewToTopocentricCoordinates", topo, CancellationToken.None);

                    if (await AwaitIfTask(taskObj).ConfigureAwait(false)) return true;
                }

                if (topo != null) {
                    object taskObj =
                        InvokeIfExistsLoose(mediator, "SlewToTopocentricCoordinates", topo, CancellationToken.None);

                    if (await AwaitIfTask(taskObj).ConfigureAwait(false)) return true;
                }
            } catch { }

            return false;
        }

        private static CancellationToken _NoCancelTokenOrDefault() => CancellationToken.None;

        private static async Task<bool> AwaitIfTask(object taskObj) {
            try {
                if (taskObj == null) return false;

                if (taskObj is Task t) {
                    await t.ConfigureAwait(false);

                    var tt = taskObj.GetType();
                    if (tt.IsGenericType && tt.GetProperty("Result") != null) {
                        try {
                            var r = tt.GetProperty("Result")?.GetValue(taskObj);
                            if (r is bool b) return b;
                        } catch { }
                    }
                    return true;
                }
                return true;
            } catch {
                return false;
            }
        }

        private static object BuildAstrometryObject(string fullTypeName, double raHours, double raDeg, double decDeg) {
            try {
                var t = FindTypeAcrossLoadedAssemblies(fullTypeName);
                if (t == null) return null;

                var ctors = t.GetConstructors(BindingFlags.Public | BindingFlags.Instance);
                foreach (var c in ctors) {
                    var ps = c.GetParameters();

                    if (ps.Length >= 2 && IsNumeric(ps[0].ParameterType) && IsNumeric(ps[1].ParameterType)) {
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

                        args[0] = ConvertNumeric(raHours, ps[0].ParameterType);
                        try {
                            var obj2 = c.Invoke(args);
                            if (obj2 != null) return obj2;
                        } catch { }
                    }
                }

                object inst = null;
                try { inst = Activator.CreateInstance(t); } catch { }
                if (inst == null) return null;

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
                var direct = Type.GetType(fullName);
                if (direct != null) return direct;

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
            object taskObj =
                InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", ra, dec) ??
                InvokeIfExistsLoose(mediator, "SlewToRaDecAsync", ra, dec) ??
                InvokeIfExistsLoose(mediator, "SlewToCoordinatesJNowAsync", ra, dec) ??
                InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", ra, dec, false) ??
                InvokeIfExistsLoose(mediator, "SlewToRaDecAsync", ra, dec, false) ??
                InvokeIfExistsLoose(mediator, "SlewToCoordinatesAsync", ra, dec, true) ??
                InvokeIfExistsLoose(mediator, "SlewToRaDecAsync", ra, dec, true);

            if (taskObj is Task t) { await t.ConfigureAwait(false); return true; }

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
                    return null;
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

        // ------------------------------
        // Plate solve value readers (ROBUST)
        // ------------------------------
        private static double ReadRaHours(object root) {
            double raH = ReadDouble(root, new[] { "RightAscensionHours", "RaHours", "RAHours" });
            if (raH != 0) return raH;

            double raDeg = ReadDouble(root, new[] { "RightAscensionDegrees", "RaDegrees", "RADegrees", "RightAscensionDeg" });
            if (raDeg != 0) return raDeg / 15.0;

            object raObj = GetAnyProp(root, new[] { "RightAscension", "RA", "Ra" });
            if (raObj != null) {
                if (TryConvertToDouble(raObj, out var v)) {
                    if (Math.Abs(v) > 24.0 + 1e-6) return v / 15.0;
                    return v;
                }
                if (raObj is string s && TryParseSexagesimalToHours(s, out var hh)) return hh;
            }

            return 0;
        }

        private static double ReadDecDeg(object root) {
            double dec = ReadDouble(root, new[] { "Declination", "DEC", "Dec", "DeclinationDegrees", "DecDegrees", "DeclinationDeg" });
            if (dec != 0) return dec;

            object decObj = GetAnyProp(root, new[] { "Declination", "DEC", "Dec" });
            if (decObj != null) {
                if (TryConvertToDouble(decObj, out var v)) return v;
                if (decObj is string s && TryParseSexagesimalToDegrees(s, out var dd)) return dd;
            }

            return 0;
        }

        private static double NormalizeAngleToDegrees(double rot, object resultRoot) {
            try {
                if (!double.IsFinite(rot)) return 0;

                double abs = Math.Abs(rot);
                if (abs > 0 && abs <= (2 * Math.PI + 0.25)) {
                    return rot * (180.0 / Math.PI);
                }

                if (abs > 7200) return rot;

                return rot;
            } catch {
                return rot;
            }
        }

        private static bool TryParseSexagesimalToHours(string s, out double hours) {
            hours = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;

            var cleaned = s.Trim()
                .Replace("h", ":").Replace("m", ":").Replace("s", "")
                .Replace("°", ":").Replace("'", ":").Replace("\"", "")
                .Replace(",", ".");
            cleaned = cleaned.Replace("  ", " ");
            cleaned = cleaned.Replace(" ", ":");

            var parts = cleaned.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return false;

            if (!double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var h)) return false;
            double m = 0, sec = 0;
            if (parts.Length > 1) double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out m);
            if (parts.Length > 2) double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out sec);

            double sign = h < 0 ? -1 : 1;
            h = Math.Abs(h);

            hours = sign * (h + (m / 60.0) + (sec / 3600.0));
            return true;
        }

        private static bool TryParseSexagesimalToDegrees(string s, out double deg) {
            deg = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;

            var cleaned = s.Trim()
                .Replace("d", ":").Replace("°", ":").Replace("'", ":").Replace("\"", "")
                .Replace(",", ".");
            cleaned = cleaned.Replace("  ", " ");
            cleaned = cleaned.Replace(" ", ":");

            var parts = cleaned.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return false;

            if (!double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return false;
            double m = 0, sec = 0;
            if (parts.Length > 1) double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out m);
            if (parts.Length > 2) double.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out sec);

            double sign = d < 0 ? -1 : 1;
            d = Math.Abs(d);

            deg = sign * (d + (m / 60.0) + (sec / 3600.0));
            return true;
        }

        private static bool TryConvertToDouble(object v, out double d) {
            d = 0;
            if (v == null) return false;

            try {
                if (v is double dd) { d = dd; return true; }
                if (v is float ff) { d = ff; return true; }
                if (v is int ii) { d = ii; return true; }
                if (v is long ll) { d = ll; return true; }
                if (v is decimal mm) { d = (double)mm; return true; }

                var t = v.GetType();

                foreach (var pn in new[] { "Degrees", "TotalDegrees", "Hours", "TotalHours", "Value" }) {
                    var pi = t.GetProperty(pn, BindingFlags.Public | BindingFlags.Instance);
                    if (pi != null) {
                        var inner = pi.GetValue(v);
                        if (inner != null && inner != v && TryConvertToDouble(inner, out d)) return true;
                    }
                }

                var s = v.ToString();
                if (!string.IsNullOrWhiteSpace(s) && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return true;
            } catch { }

            return false;
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

                    if (TryConvertToDouble(v, out var d)) return d;

                    if (v is string s && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)) return parsed;
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

        private static object GetAnyProp(object root, string[] propNames) {
            if (root == null) return null;
            try {
                var t = root.GetType();
                foreach (var p in propNames) {
                    var pi = t.GetProperty(p, BindingFlags.Public | BindingFlags.Instance);
                    if (pi == null) continue;
                    var v = pi.GetValue(root);
                    if (v != null) return v;
                }
            } catch { }
            return null;
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
                    if (TryConvertToDouble(v, out var d)) return d;
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

        // =========================================================
        // NEW: NINA log tailer fallback (no NINA API needed)
        // =========================================================
        private readonly struct NinaSolveLine {
            public NinaSolveLine(double raHours, double decDeg, double paDeg) {
                RaHours = raHours;
                DecDeg = decDeg;
                PositionAngleDeg = paDeg;
            }
            public double RaHours { get; }
            public double DecDeg { get; }
            public double PositionAngleDeg { get; }
        }

        private sealed class NinaLogTailer : IDisposable {
            private readonly CancellationToken _ct;
            private readonly Action<NinaSolveLine> _onSolve;
            private readonly object _gate = new object();

            private Task _task;
            private string _currentFile;
            private long _pos;

            // Matches example:
            // "...Platesolve successful: Coordinates: RA: 05:34:12; Dec: -05° 28' 12\"; Epoch: J2000 - Position Angle: 23.5030"
            private static readonly Regex SolveOk =
                new Regex(@"Platesolve successful:\s*Coordinates:\s*RA:\s*(?<ra>\d{1,2}:\d{2}:\d{2})\s*;\s*Dec:\s*(?<dec>[-+−]?\d{1,2})\D+(?<dm>\d{1,2})\D+(?<ds>\d{1,2})\D+.*?Position Angle:\s*(?<pa>[-+−]?\d+(\.\d+)?)",
                          RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

            public NinaLogTailer(CancellationToken ct, Action<NinaSolveLine> onSolve) {
                _ct = ct;
                _onSolve = onSolve;
            }

            public void Start() {
                _task = Task.Run(Loop, _ct);
            }

            private async Task Loop() {
                while (!_ct.IsCancellationRequested) {
                    try {
                        EnsureCurrentLogFile();

                        if (!string.IsNullOrWhiteSpace(_currentFile) && File.Exists(_currentFile)) {
                            using (var fs = new FileStream(_currentFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            using (var sr = new StreamReader(fs, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true)) {
                                if (_pos > fs.Length) _pos = 0; // log rotated/truncated
                                fs.Seek(_pos, SeekOrigin.Begin);
                                string line;
                                while ((line = await sr.ReadLineAsync().ConfigureAwait(false)) != null) {
                                    _pos = fs.Position;
                                    TryParseSolveLine(line);
                                }
                            }
                        }
                    } catch {
                        // ignore; retry
                    }

                    try { await Task.Delay(350, _ct).ConfigureAwait(false); } catch { }
                }
            }

            private void EnsureCurrentLogFile() {
                // NINA logs typically in: %LOCALAPPDATA%\NINA\Logs\*.log
                // If that folder doesn't exist, we do nothing.
                try {
                    var logsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NINA", "Logs");
                    if (!Directory.Exists(logsDir)) return;

                    var newest = Directory.EnumerateFiles(logsDir, "*.log", SearchOption.TopDirectoryOnly)
                        .Select(p => new FileInfo(p))
                        .OrderByDescending(fi => fi.LastWriteTimeUtc)
                        .FirstOrDefault();

                    if (newest == null) return;

                    lock (_gate) {
                        if (!string.Equals(_currentFile, newest.FullName, StringComparison.OrdinalIgnoreCase)) {
                            _currentFile = newest.FullName;
                            _pos = 0; // start from beginning (safe); logs are not huge
                        }
                    }
                } catch { }
            }

            private void TryParseSolveLine(string line) {
                try {
                    if (string.IsNullOrWhiteSpace(line)) return;
                    var m = SolveOk.Match(line);
                    if (!m.Success) return;

                    var raStr = m.Groups["ra"].Value; // HH:MM:SS
                    if (!TryParseRaHmsToHours(raStr, out var raHours)) return;

                    int d = ParseInt(m.Groups["dec"].Value);
                    int dm = ParseInt(m.Groups["dm"].Value);
                    int ds = ParseInt(m.Groups["ds"].Value);
                    double sign = d < 0 ? -1.0 : 1.0;
                    double decDeg = sign * (Math.Abs(d) + (dm / 60.0) + (ds / 3600.0));

                    var paStr = m.Groups["pa"].Value.Replace('−', '-');
                    if (!double.TryParse(paStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var paDeg)) paDeg = 0;

                    _onSolve?.Invoke(new NinaSolveLine(raHours, decDeg, paDeg));
                } catch { }
            }

            private static int ParseInt(string s) {
                if (string.IsNullOrWhiteSpace(s)) return 0;
                s = s.Replace('−', '-');
                int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v);
                return v;
            }

            private static bool TryParseRaHmsToHours(string hms, out double hours) {
                hours = 0;
                if (string.IsNullOrWhiteSpace(hms)) return false;
                var parts = hms.Split(':');
                if (parts.Length != 3) return false;

                if (!double.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out var h)) return false;
                if (!double.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var m)) return false;
                if (!double.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out var s)) return false;

                hours = h + (m / 60.0) + (s / 3600.0);
                // Normalize
                while (hours < 0) hours += 24;
                while (hours >= 24) hours -= 24;
                return true;
            }

            public void Dispose() {
                // nothing special; cancellation handled externally
            }
        }
        // =========================================================
    }
}