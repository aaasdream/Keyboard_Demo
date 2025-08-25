// Form1.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks; // 引入 Task 命名空間
using System.Windows.Forms;

namespace TouchKeyBoard
{
    public partial class Form1 : Form
    {
        #region 狀態管理與變數

        private enum AppState { Normal, IncomingCall }
        private AppState _currentState = AppState.Normal;

        private readonly bool _debug = false;
        private DebugWindowLogger _logger;
        private Button[] _buttons = new Button[10];
        private string _currentFocusedAppName = "";
        private IntPtr _lastActiveWindow = IntPtr.Zero;
        private readonly string _configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
        private Dictionary<string, List<ShortcutKey>> _appShortcuts;
        private System.Windows.Forms.Timer _focusCheckTimer = new();

        private TeamsDetector _teamsDetector;
        private IntPtr _teamsCallWindowHandle = IntPtr.Zero;

        #endregion

        #region 表單初始化與核心修改

        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
        }

        private void InitializeCustomComponents()
        {
            this.Text = "TouchKeyBoard";
            this.TopMost = true;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.ShowInTaskbar = false;

            if (_debug)
            {
                _logger = new DebugWindowLogger(new Point(this.Location.X + this.Width + 10, this.Location.Y));
            }

            LogDebugMessage($"調試模式已啟動 - {DateTime.Now:HH:mm:ss}");

            var configManager = new ConfigManager(_configFilePath, LogDebugMessage);
            _appShortcuts = configManager.LoadShortcuts();

            CreateButtons();

            _teamsDetector = new TeamsDetector(LogDebugMessage);
            _teamsDetector.CallDetected += OnTeamsCallDetected;
            _teamsDetector.CallEnded += OnTeamsCallEnded;
            _teamsDetector.Start();

            _focusCheckTimer.Interval = 500;
            _focusCheckTimer.Tick += CheckFocusedWindow;
            _focusCheckTimer.Start();

            LogDebugMessage("正在監控視窗變化...");
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= NativeMethods.WS_EX_NOACTIVATE;
                return cp;
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _teamsDetector?.Dispose();
            _logger?.Dispose();
            base.OnFormClosed(e);
        }

        #endregion

        #region 調試功能
        private void LogDebugMessage(string message)
        {
            _logger?.Log(message);
        }
        #endregion

        #region 焦點監控與狀態切換

        private void CheckFocusedWindow(object sender, EventArgs e)
        {
            NativeMethods.SetWindowPos(this.Handle, NativeMethods.HWND_TOPMOST, 0, 0, 0, 0, NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);

            if (_currentState == AppState.IncomingCall)
            {
                if (_teamsCallWindowHandle == IntPtr.Zero || !NativeMethods.IsWindowVisible(_teamsCallWindowHandle))
                {
                    LogDebugMessage("[安全網] 偵測到來電通知視窗已消失，強制恢復 Normal 狀態。");
                    ChangeState(AppState.Normal, IntPtr.Zero, null);
                }
                return;
            }

            IntPtr currentWindow = NativeMethods.GetForegroundWindow();
            if (currentWindow == this.Handle) return; // 忽略自己

            if (FocusManager.IsTeamsInCallWindow(currentWindow))
            {
                // ★★★ 改善點 #1：當定時器偵測到通話視窗時，也更新 _lastActiveWindow ★★★
                // 這樣即使用戶手動切換到其他視窗再切回來，也能保持正確的目標
                if (_currentFocusedAppName != "TeamsInCall" || _lastActiveWindow != currentWindow)
                {
                    LogDebugMessage($"[Focus Check] 偵測到 Teams 通話視窗 (Handle: {currentWindow})，更新狀態。");
                    _currentFocusedAppName = "TeamsInCall";
                    _lastActiveWindow = currentWindow; // 確保目標視窗為當前的通話視窗
                    SetupInCallButtons();
                }
            }
            else
            {
                string appName = FocusManager.GetProcessNameFromHwnd(currentWindow);
                if (!string.IsNullOrEmpty(appName) && appName != "TouchKeyBoard")
                {
                    _lastActiveWindow = currentWindow;
                    if (appName != _currentFocusedAppName)
                    {
                        _currentFocusedAppName = appName;
                        SetupNormalButtons();
                    }
                }
            }
        }

        private void ChangeState(AppState newState, IntPtr windowHandle, List<string> teamsActions)
        {
            if (_currentState == newState) return;

            LogDebugMessage($"--- 狀態變更: {_currentState} -> {newState} ---");
            _currentState = newState;

            _teamsCallWindowHandle = (newState == AppState.IncomingCall) ? windowHandle : IntPtr.Zero;

            this.Invoke(new Action(() => UpdateButtonsForState(teamsActions)));
        }
        #endregion

        #region Teams 事件處理
        private void OnTeamsCallDetected(IntPtr hwnd, List<string> actions)
        {
            if (_currentState == AppState.IncomingCall) return; // 避免重複進入
            ChangeState(AppState.IncomingCall, hwnd, actions);
        }

        private void OnTeamsCallEnded(IntPtr hwnd)
        {
            if (hwnd == _teamsCallWindowHandle)
            {
                ChangeState(AppState.Normal, IntPtr.Zero, null);
            }
        }
        #endregion

        #region UI 創建與更新
        private void UpdateButtonsForState(List<string> teamsActions)
        {
            this.Text = $"TouchKeyBoard - {_currentState}";
            switch (_currentState)
            {
                case AppState.IncomingCall:
                    SetupIncomingCallButtons(teamsActions);
                    break;
                case AppState.Normal:
                default:
                    // 恢復Normal狀態後，立即重新檢查一次當前視窗，確保UI正確
                    _currentFocusedAppName = "";
                    CheckFocusedWindow(null, null);
                    break;
            }
        }

        private void SetupNormalButtons()
        {
            this.Text = $"TouchKeyBoard - {_currentFocusedAppName}";
            foreach (Button btn in _buttons)
            {
                btn.Text = ""; btn.Tag = null; btn.Enabled = false;
                btn.BackColor = SystemColors.Control; btn.ForeColor = SystemColors.ControlText;
            }

            var appKey = _appShortcuts.Keys.FirstOrDefault(k => _currentFocusedAppName.ToLower().Contains(k.ToLower()));
            if (appKey == null) return;

            var shortcuts = _appShortcuts[appKey];
            for (int i = 0; i < shortcuts.Count && i < _buttons.Length; i++)
            {
                _buttons[i].Text = $"{shortcuts[i].DisplayName}\n{shortcuts[i].KeyCombination}";
                _buttons[i].Tag = shortcuts[i];
                _buttons[i].Enabled = true;
            }
        }

        private void SetupIncomingCallButtons(List<string> actions)
        {
            this.Text = "TouchKeyBoard - Incoming Call";
            foreach (Button btn in _buttons)
            {
                btn.Text = ""; btn.Tag = null; btn.Enabled = false;
                btn.BackColor = SystemColors.Control; btn.ForeColor = SystemColors.ControlText;
            }
            int buttonIndex = 0;
            if (actions.Contains("Video"))
            {
                _buttons[buttonIndex].Text = "接聽視訊";
                _buttons[buttonIndex].Tag = new ShortcutKey { DisplayName = "接聽視訊", KeyCombination = "Ctrl+Shift+A", SendKeysFormat = "^+a" };
                _buttons[buttonIndex].Enabled = true;
                _buttons[buttonIndex].BackColor = Color.FromArgb(0, 120, 212);
                _buttons[buttonIndex].ForeColor = Color.White;
                buttonIndex++;
            }
            if (actions.Contains("Audio"))
            {
                _buttons[buttonIndex].Text = "接聽語音";
                _buttons[buttonIndex].Tag = new ShortcutKey { DisplayName = "接聽語音", KeyCombination = "Ctrl+Shift+S", SendKeysFormat = "^+s" };
                _buttons[buttonIndex].Enabled = true;
                _buttons[buttonIndex].BackColor = Color.FromArgb(4, 153, 114);
                _buttons[buttonIndex].ForeColor = Color.White;
                buttonIndex++;
            }
            if (actions.Contains("Decline"))
            {
                _buttons[buttonIndex].Text = "拒絕";
                _buttons[buttonIndex].Tag = new ShortcutKey { DisplayName = "拒絕", KeyCombination = "Ctrl+Shift+D", SendKeysFormat = "^+d" };
                _buttons[buttonIndex].Enabled = true;
                _buttons[buttonIndex].BackColor = Color.FromArgb(217, 48, 48);
                _buttons[buttonIndex].ForeColor = Color.White;
            }
        }

        private void SetupInCallButtons()
        {
            this.Text = "TouchKeyBoard - In Call";
            foreach (Button btn in _buttons)
            {
                btn.Text = ""; btn.Tag = null; btn.Enabled = false;
                btn.BackColor = SystemColors.Control; btn.ForeColor = SystemColors.ControlText;
            }
            _buttons[0].Text = "靜音切換";
            _buttons[0].Tag = new ShortcutKey { DisplayName = "靜音切換", KeyCombination = "Ctrl+Shift+M", SendKeysFormat = "^+m" };
            _buttons[0].Enabled = true;
            _buttons[0].BackColor = Color.FromArgb(0, 120, 212);
            _buttons[0].ForeColor = Color.White;

            _buttons[1].Text = "視訊切換";
            _buttons[1].Tag = new ShortcutKey { DisplayName = "視訊切換", KeyCombination = "Ctrl+Shift+O", SendKeysFormat = "^+o" };
            _buttons[1].Enabled = true;
            _buttons[1].BackColor = Color.DarkSlateBlue;
            _buttons[1].ForeColor = Color.White;

            _buttons[2].Text = "掛斷";
            _buttons[2].Tag = new ShortcutKey { DisplayName = "掛斷", KeyCombination = "Ctrl+Shift+H", SendKeysFormat = "^+h" };
            _buttons[2].Enabled = true;
            _buttons[2].BackColor = Color.FromArgb(217, 48, 48);
            _buttons[2].ForeColor = Color.White;
        }

        private void Button_Click(object sender, EventArgs e)
        {
            if (sender is not Button button || button.Tag is not ShortcutKey shortcut) return;

            LogDebugMessage($">>> 執行快捷鍵: {shortcut.DisplayName} ({shortcut.SendKeysFormat})");

            IntPtr targetWindow = (_currentState == AppState.IncomingCall) ? _teamsCallWindowHandle : _lastActiveWindow;
            LogDebugMessage($" -> 模式: {_currentState}, 目標視窗: {targetWindow} ({WindowInspector.GetWindowDebugInfo(targetWindow)})");

            if (targetWindow == IntPtr.Zero)
            {
                LogDebugMessage("✗ 目標窗口句柄為空，無法發送按鍵。");
                return;
            }

            try
            {
                // 先將目標視窗設為前景
                bool switchResult = NativeMethods.SetForegroundWindow(targetWindow);
                LogDebugMessage($"切換窗口結果: {(switchResult ? "成功" : "失敗")}");

                // 即使切換失敗，仍嘗試發送，因為有時視窗已在前景但API返回false
                System.Threading.Thread.Sleep(100); // 等待一下確保焦點穩定
                SendKeys.SendWait(shortcut.SendKeysFormat);
                LogDebugMessage($"✓ 快捷鍵已發送");

                if (_currentState == AppState.IncomingCall)
                {
                    HandleIncomingCallAction(shortcut, targetWindow);
                }
            }
            catch (Exception ex)
            {
                LogDebugMessage($"✗ 發送快捷鍵失敗: {ex.Message}");
            }
        }

        // ★★★ 核心修改點 ★★★
        // 修改此方法，使其在找到新視窗後，透過 Invoke 更新主 UI 執行緒的狀態
        private async void FocusTeamsInCallWindowAsync(IntPtr oldNotificationHwnd)
        {
            LogDebugMessage("-> 啟動非同步任務：搜尋新的 Teams 通話視窗...");

            await Task.Run(async () =>
            {
                // 在最多 5 秒內，每 250 毫秒嘗試一次
                for (int i = 0; i < 20; i++)
                {
                    Process[] teamsProcesses = Process.GetProcessesByName("ms-teams");
                    foreach (Process p in teamsProcesses)
                    {
                        IntPtr candidateHwnd = p.MainWindowHandle;

                        if (candidateHwnd != IntPtr.Zero &&
                            candidateHwnd != oldNotificationHwnd &&
                            FocusManager.IsTeamsInCallWindow(candidateHwnd))
                        {
                            LogDebugMessage($"  ✓ 找到新的通話視窗 Handle: {candidateHwnd}");

                            // ★★★ 核心：在 UI 執行緒上更新狀態並設定前景視窗 ★★★
                            this.Invoke(new Action(() => {
                                LogDebugMessage("  -> 正在 UI 執行緒上更新主表單狀態...");
                                NativeMethods.SetForegroundWindow(candidateHwnd);
                                _lastActiveWindow = candidateHwnd; // <-- 關鍵修復：更新最後活動視窗
                                _currentFocusedAppName = "TeamsInCall"; // <-- 關鍵修復：更新當前應用程式名稱
                                SetupInCallButtons(); // <-- 關鍵修復：立即刷新按鈕為「通話中」模式
                                LogDebugMessage($"  -> 主狀態已成功更新: _lastActiveWindow = {candidateHwnd}, AppName = TeamsInCall");
                            }));

                            LogDebugMessage("  ✓✓✓ 成功觸發焦點切換與狀態更新。搜尋任務結束。");
                            return; // 找到後立刻結束任務
                        }
                    }
                    await Task.Delay(250); // 等待一下再進行下一次搜尋
                }
                LogDebugMessage("  ✗ 搜尋超時，未能找到新的 Teams 通話視窗。");
            });
        }

        private void HandleIncomingCallAction(ShortcutKey shortcut, IntPtr windowToClose)
        {
            string debugInfo = WindowInspector.GetWindowDebugInfo(windowToClose);
            LogDebugMessage($"準備對以下視窗執行 '{shortcut.DisplayName}' 操作：");
            LogDebugMessage(debugInfo);

            bool wasActionTaken = false;
            if (shortcut.DisplayName.Contains("接聽"))
            {
                var acceptNames = new List<string> { "Accept with video", "接聽視訊", "Accept with audio", "接聽語音", "Accept", "接聽" };
                wasActionTaken = UIAutomationController.InvokeButtonByName(windowToClose, acceptNames, LogDebugMessage);

                // 只有在成功點擊後，才啟動焦點搜尋任務
                if (wasActionTaken)
                {
                    FocusTeamsInCallWindowAsync(windowToClose);
                }
            }
            else if (shortcut.DisplayName.Contains("拒絕"))
            {
                var declineNames = new List<string> { "Decline", "拒絕" };
                wasActionTaken = UIAutomationController.InvokeButtonByName(windowToClose, declineNames, LogDebugMessage);
            }

            // 無論成功與否，都應離開來電狀態
            ChangeState(AppState.Normal, IntPtr.Zero, null);
        }

        private void CreateButtons()
        {
            int buttonHeight = ClientSize.Height / 10;
            for (int i = 0; i < 10; i++)
            {
                _buttons[i] = new Button
                {
                    Width = ClientSize.Width,
                    Height = buttonHeight,
                    Location = new Point(0, i * buttonHeight),
                    Enabled = false,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                    TabStop = false,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold)
                };
                _buttons[i].FlatAppearance.BorderSize = 1;
                _buttons[i].Click += Button_Click;
                Controls.Add(_buttons[i]);
            }
        }
        #endregion
    }
}