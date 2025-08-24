using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows.Automation; // 需要參考 UIAutomationClient 和 UIAutomationTypes

namespace TouchKeyBoard
{
    public partial class Form1 : Form
    {
        #region Win32 API 導入與常數

        // --- 視窗定位與樣式設定 ---
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);
        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        // --- 用於發送關閉訊息的 API ---
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        private const uint WM_CLOSE = 0x0010;


        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);
        private const uint EVENT_OBJECT_SHOW = 0x8002;
        private const uint EVENT_OBJECT_HIDE = 0x8003;
        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const string TEAMS_CALL_WINDOW_CLASS = "TeamsWebView";
        private const string TEAMS_PROCESS_NAME = "ms-teams";

        #endregion

        #region 狀態管理與變數

        private enum AppState
        {
            Normal,
            IncomingCall,
            InCall
        }
        private AppState currentState = AppState.Normal;

        private DateTime inCallStateChangeTime;
        private const int IN_CALL_GRACE_PERIOD_SECONDS = 5;

        private bool debug = true;
        private Form debugWindow;
        private TextBox debugTextBox;
        private Button[] buttons = new Button[10];
        private string currentFocusedAppName = "";
        private IntPtr lastActiveWindow = IntPtr.Zero;
        private readonly string configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
        private Dictionary<string, List<ShortcutKey>> appShortcuts = new();
        private System.Windows.Forms.Timer focusCheckTimer = new();
        private WinEventDelegate teamsCallDelegate;
        private IntPtr teamsEventHook = IntPtr.Zero;
        private IntPtr teamsCallWindowHandle = IntPtr.Zero;

        #endregion

        #region 表單初始化與核心修改

        public Form1()
        {
            InitializeComponent();

            this.Text = "TouchKeyBoard";
            this.TopMost = true;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.ShowInTaskbar = false;

            Directory.CreateDirectory(Path.GetDirectoryName(configFilePath));

            LoadConfig();
            CreateButtons();
            InitializeTeamsHook();

            if (debug)
            {
                CreateDebugWindow();
            }

            focusCheckTimer.Interval = 500;
            focusCheckTimer.Tick += CheckFocusedWindow;
            focusCheckTimer.Start();

            if (debug)
            {
                LogDebugMessage($"調試模式已啟動 - {DateTime.Now:HH:mm:ss}");
                LogDebugMessage($"設定檔路徑: {configFilePath}");
                LogDebugMessage("正在監控視窗變化...");
                LogDebugMessage("Teams 來電監聽已啟動。");
            }

            this.FormClosed += (s, e) => {
                if (teamsEventHook != IntPtr.Zero)
                {
                    UnhookWinEvent(teamsEventHook);
                    if (debug) LogDebugMessage("Teams 來電監聽掛鉤已卸載。");
                }
                debugWindow?.Close();
            };
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WS_EX_NOACTIVATE;
                return cp;
            }
        }

        #endregion

        #region 調試功能

        private void CreateDebugWindow()
        {
            debugWindow = new Form
            {
                Text = "TouchKeyBoard Debug Window",
                Size = new Size(500, 400),
                Location = new Point(this.Location.X + this.Width + 10, this.Location.Y),
                TopMost = true,
                FormBorderStyle = FormBorderStyle.SizableToolWindow,
                BackColor = Color.Black
            };

            debugTextBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                BackColor = Color.Black,
                ForeColor = Color.Lime,
                Font = new Font("Consolas", 10, FontStyle.Regular),
                ReadOnly = true,
                WordWrap = true
            };

            debugWindow.Controls.Add(debugTextBox);
            debugWindow.Show();
        }

        private void LogDebugMessage(string message)
        {
            if (!debug || debugTextBox == null) return;

            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string logMessage = $"[{timestamp}] {message}";

            if (debugTextBox.InvokeRequired)
            {
                debugTextBox.Invoke(new Action(() =>
                {
                    debugTextBox.AppendText(logMessage + Environment.NewLine);
                    debugTextBox.SelectionStart = debugTextBox.Text.Length;
                    debugTextBox.ScrollToCaret();
                }));
            }
            else
            {
                debugTextBox.AppendText(logMessage + Environment.NewLine);
                debugTextBox.SelectionStart = debugTextBox.Text.Length;
                debugTextBox.ScrollToCaret();
            }
        }
        #endregion

        #region 設定檔處理 (JSON)

        private void LoadConfig()
        {
            try
            {
                if (File.Exists(configFilePath))
                {
                    if (debug) LogDebugMessage("找到設定檔，正在讀取...");
                    string json = File.ReadAllText(configFilePath);
                    var loadedSettings = JsonSerializer.Deserialize<Dictionary<string, List<ShortcutKey>>>(json);

                    if (loadedSettings != null)
                    {
                        appShortcuts = loadedSettings;
                        if (debug) LogDebugMessage($"✓ 設定檔載入成功，共 {appShortcuts.Count} 個應用程式設定。");
                    }
                    else
                    {
                        if (debug) LogDebugMessage("✗ 設定檔為空或格式錯誤，將使用空設定。");
                        appShortcuts = new Dictionary<string, List<ShortcutKey>>();
                    }
                }
                else
                {
                    if (debug) LogDebugMessage("設定檔不存在，將創建並載入預設設定檔。");
                    CreateDefaultConfigFile();
                    LoadConfig();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加載設定檔時出錯: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (debug) LogDebugMessage($"✗ 載入設定檔失敗: {ex.Message}");
                appShortcuts = new Dictionary<string, List<ShortcutKey>>();
            }
        }
        private void CreateDefaultConfigFile()
        {
            try
            {
                string defaultConfigJson = GetDefaultConfigJsonString();
                File.WriteAllText(configFilePath, defaultConfigJson);
                if (debug) LogDebugMessage("✓ 預設設定檔已創建成功。");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"創建預設設定檔時出錯: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (debug) LogDebugMessage($"✗ 創建預設設定檔失敗: {ex.Message}");
            }
        }
        private string GetDefaultConfigJsonString()
        {
            return """
            {
              "mspaint": [
                { "DisplayName": "新增", "KeyCombination": "Ctrl+N", "SendKeysFormat": "^n", "Description": "" },
                { "DisplayName": "開啟", "KeyCombination": "Ctrl+O", "SendKeysFormat": "^o", "Description": "" },
                { "DisplayName": "儲存", "KeyCombination": "Ctrl+S", "SendKeysFormat": "^s", "Description": "" },
                { "DisplayName": "另存新檔", "KeyCombination": "F12", "SendKeysFormat": "{F12}", "Description": "" },
                { "DisplayName": "復原", "KeyCombination": "Ctrl+Z", "SendKeysFormat": "^z", "Description": "" },
                { "DisplayName": "重做", "KeyCombination": "Ctrl+Y", "SendKeysFormat": "^y", "Description": "" },
                { "DisplayName": "全選", "KeyCombination": "Ctrl+A", "SendKeysFormat": "^a", "Description": "" },
                { "DisplayName": "複製", "KeyCombination": "Ctrl+C", "SendKeysFormat": "^c", "Description": "" },
                { "DisplayName": "貼上", "KeyCombination": "Ctrl+V", "SendKeysFormat": "^v", "Description": "" },
                { "DisplayName": "放大", "KeyCombination": "Ctrl+Plus", "SendKeysFormat": "^{ADD}", "Description": "" }
              ],
              "notepad": [
                { "DisplayName": "新增", "KeyCombination": "Ctrl+N", "SendKeysFormat": "^n", "Description": "" },
                { "DisplayName": "開啟", "KeyCombination": "Ctrl+O", "SendKeysFormat": "^o", "Description": "" },
                { "DisplayName": "儲存", "KeyCombination": "Ctrl+S", "SendKeysFormat": "^s", "Description": "" },
                { "DisplayName": "另存新檔", "KeyCombination": "Ctrl+Shift+S", "SendKeysFormat": "^+s", "Description": "" },
                { "DisplayName": "尋找", "KeyCombination": "Ctrl+F", "SendKeysFormat": "^f", "Description": "" },
                { "DisplayName": "取代", "KeyCombination": "Ctrl+H", "SendKeysFormat": "^h", "Description": "" },
                { "DisplayName": "全選", "KeyCombination": "Ctrl+A", "SendKeysFormat": "^a", "Description": "" },
                { "DisplayName": "復原", "KeyCombination": "Ctrl+Z", "SendKeysFormat": "^z", "Description": "" },
                { "DisplayName": "複製", "KeyCombination": "Ctrl+C", "SendKeysFormat": "^c", "Description": "" },
                { "DisplayName": "時間/日期", "KeyCombination": "F5", "SendKeysFormat": "{F5}", "Description": "" }
              ],
              "Code": [
                { "DisplayName": "儲存", "KeyCombination": "Ctrl+S", "SendKeysFormat": "^s", "Description": "" },
                { "DisplayName": "開啟檔案", "KeyCombination": "Ctrl+O", "SendKeysFormat": "^o", "Description": "" },
                { "DisplayName": "新增檔案", "KeyCombination": "Ctrl+N", "SendKeysFormat": "^n", "Description": "" },
                { "DisplayName": "尋找", "KeyCombination": "Ctrl+F", "SendKeysFormat": "^f", "Description": "" },
                { "DisplayName": "取代", "KeyCombination": "Ctrl+H", "SendKeysFormat": "^h", "Description": "" },
                { "DisplayName": "註解", "KeyCombination": "Ctrl+/", "SendKeysFormat": "^/", "Description": "" },
                { "DisplayName": "復原", "KeyCombination": "Ctrl+Z", "SendKeysFormat": "^z", "Description": "" },
                { "DisplayName": "重做", "KeyCombination": "Ctrl+Y", "SendKeysFormat": "^y", "Description": "" },
                { "DisplayName": "執行", "KeyCombination": "F5", "SendKeysFormat": "{F5}", "Description": "" },
                { "DisplayName": "格式化", "KeyCombination": "Shift+Alt+F", "SendKeysFormat": "+%f", "Description": "" }
              ],
              "chrome": [
                { "DisplayName": "新分頁", "KeyCombination": "Ctrl+T", "SendKeysFormat": "^t", "Description": "" },
                { "DisplayName": "關閉分頁", "KeyCombination": "Ctrl+W", "SendKeysFormat": "^w", "Description": "" },
                { "DisplayName": "重新整理", "KeyCombination": "F5", "SendKeysFormat": "{F5}", "Description": "" },
                { "DisplayName": "上一頁", "KeyCombination": "Alt+Left", "SendKeysFormat": "%{LEFT}", "Description": "" },
                { "DisplayName": "下一頁", "KeyCombination": "Alt+Right", "SendKeysFormat": "%{RIGHT}", "Description": "" },
                { "DisplayName": "書籤", "KeyCombination": "Ctrl+D", "SendKeysFormat": "^d", "Description": "" },
                { "DisplayName": "開發者工具", "KeyCombination": "F12", "SendKeysFormat": "{F12}", "Description": "" },
                { "DisplayName": "無痕模式", "KeyCombination": "Ctrl+Shift+N", "SendKeysFormat": "^+n", "Description": "" },
                { "DisplayName": "尋找", "KeyCombination": "Ctrl+F", "SendKeysFormat": "^f", "Description": "" },
                { "DisplayName": "下載", "KeyCombination": "Ctrl+J", "SendKeysFormat": "^j", "Description": "" }
              ]
            }
            """;
        }
        #endregion

        #region Teams 來電監聽

        private void InitializeTeamsHook()
        {
            teamsCallDelegate = new WinEventDelegate(WinEventProc);
            teamsEventHook = SetWinEventHook(EVENT_OBJECT_SHOW, EVENT_OBJECT_HIDE, IntPtr.Zero, teamsCallDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);
        }

        private void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (idObject != 0 || hwnd == IntPtr.Zero) return;
            if (currentState == AppState.InCall && eventType == EVENT_OBJECT_SHOW) return;

            StringBuilder className = new StringBuilder(256);
            GetClassName(hwnd, className, className.Capacity);

            if (className.ToString() == TEAMS_CALL_WINDOW_CLASS)
            {
                if (eventType == EVENT_OBJECT_SHOW)
                {
                    var availableActions = DetectTeamsCallActions(hwnd);
                    if (availableActions.Count > 0)
                    {
                        ChangeState(AppState.IncomingCall, hwnd, availableActions);
                    }
                }
                else if (eventType == EVENT_OBJECT_HIDE)
                {
                    if (hwnd == teamsCallWindowHandle)
                    {
                        ChangeState(AppState.Normal, IntPtr.Zero, null);
                    }
                }
            }
        }

        private List<string> DetectTeamsCallActions(IntPtr hwnd)
        {
            var availableActions = new List<string>();
            if (!GetWindowRect(hwnd, out RECT rect) || (rect.Right - rect.Left) <= 300) return availableActions;

            if (debug) LogDebugMessage($" -> 尺寸符合. 開始 UIA 掃描...");
            try
            {
                AutomationElement rootElement = AutomationElement.FromHandle(hwnd);
                if (rootElement == null) return availableActions;

                var buttonCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                var allButtons = rootElement.FindAll(TreeScope.Descendants, buttonCondition);

                string[] videoAcceptNames = { "Accept with video", "接聽視訊" };
                string[] audioAcceptNames = { "Accept with audio", "接聽語音", "Accept", "接聽" };
                string[] declineNames = { "Decline", "拒絕" };

                if (debug) LogDebugMessage($"   - 掃描到 {allButtons.Count} 個按鈕元件。");
                foreach (AutomationElement button in allButtons)
                {
                    try
                    {
                        string name = button.Current.Name;
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        if (debug) LogDebugMessage($"   - 找到按鈕: '{name}'");

                        if (videoAcceptNames.Any(x => name.Contains(x, StringComparison.OrdinalIgnoreCase)) && !availableActions.Contains("Video"))
                            availableActions.Add("Video");
                        if (audioAcceptNames.Any(x => name.Contains(x, StringComparison.OrdinalIgnoreCase)) && !availableActions.Contains("Audio"))
                            availableActions.Add("Audio");
                        if (declineNames.Any(x => name.Contains(x, StringComparison.OrdinalIgnoreCase)) && !availableActions.Contains("Decline"))
                            availableActions.Add("Decline");
                    }
                    catch (ElementNotAvailableException) { }
                }
            }
            catch (Exception ex)
            {
                if (debug) LogDebugMessage($"   - UIA 掃描出錯: {ex.Message}");
            }
            if (debug) LogDebugMessage($" -> 掃描完成. 偵測到的操作: [{string.Join(", ", availableActions)}]");
            return availableActions;
        }

        #endregion

        #region 焦點監控與狀態切換

        // ▼▼▼ 核心修正：完全重構 CheckFocusedWindow 的邏輯 ▼▼▼
        private void CheckFocusedWindow(object sender, EventArgs e)
        {
            SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

            IntPtr currentWindow = GetForegroundWindow();
            if (currentWindow == this.Handle || (debugWindow != null && currentWindow == debugWindow.Handle))
            {
                return;
            }

            // 狀態機不應該在有來電時被 timer 影響，WinEventProc 優先
            if (currentState == AppState.IncomingCall)
            {
                return;
            }

            // 優先級 1: 檢查當前視窗是否為通話視窗
            if (IsTeamsCallWindow(currentWindow))
            {
                // 如果是通話視窗，我們的狀態必須是 InCall
                if (currentState != AppState.InCall)
                {
                    LogDebugMessage(" -> 偵測到焦點在通話視窗上，但目前狀態不是 InCall。強制切換！");
                    ChangeState(AppState.InCall, currentWindow, null);
                }
                else
                {
                    // 如果狀態已經是 InCall，只需確保句柄是最新的
                    teamsCallWindowHandle = currentWindow;
                }
                return; // 處理完畢，結束本次檢查
            }

            // 優先級 2: 如果不是通話視窗，再處理其他情況
            if (currentState == AppState.InCall)
            {
                // 如果我們以為在通話中，但當前視窗已不是通話視窗，則準備切換回 Normal
                bool isInGracePeriod = (DateTime.Now - inCallStateChangeTime).TotalSeconds < IN_CALL_GRACE_PERIOD_SECONDS;
                if (isInGracePeriod)
                {
                    LogDebugMessage($" -> UIA 檢查失敗，但仍在 {IN_CALL_GRACE_PERIOD_SECONDS} 秒寬限期內。忽略並保持 InCall 狀態。");
                    return; // 在寬限期內，忽略這次的失敗
                }
                else
                {
                    string appName = GetProcessNameFromHwnd(currentWindow);
                    LogDebugMessage($" -> 焦點已離開通話視窗 (目前在 '{appName}') 且寬限期已過。恢復為 Normal 狀態。");
                    ChangeState(AppState.Normal, IntPtr.Zero, null);
                }
            }
            else if (currentState == AppState.Normal)
            {
                // 執行標準的應用程式快捷鍵邏輯
                string appName = GetProcessNameFromHwnd(currentWindow);
                if (!string.IsNullOrEmpty(appName))
                {
                    if (appName != "TouchKeyBoard")
                    {
                        lastActiveWindow = currentWindow;
                    }
                    if (appName != currentFocusedAppName)
                    {
                        currentFocusedAppName = appName;
                        UpdateButtonsForApp(null);
                    }
                }
            }
        }

        private bool IsTeamsCallWindow(IntPtr hwnd)
        {
            string processName = GetProcessNameFromHwnd(hwnd);
            if (!processName.Equals(TEAMS_PROCESS_NAME, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            StringBuilder titleBuilder = new StringBuilder(256);
            GetWindowText(hwnd, titleBuilder, titleBuilder.Capacity);
            if (debug) LogDebugMessage($" -> UIA 視窗檢查: 開始掃描句柄 {hwnd}, 標題: '{titleBuilder}'...");

            string[] inCallButtonKeywords = {
                "Mute", "Unmute", "靜音",
                "Camera", "攝影機", "Turn camera on", "Turn camera off",
                "Leave", "Hang up", "離開", "掛斷",
                "Share", "共用"
            };

            try
            {
                AutomationElement rootElement = AutomationElement.FromHandle(hwnd);
                if (rootElement == null) return false;

                var buttonCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                var allButtons = rootElement.FindAll(TreeScope.Descendants, buttonCondition);

                if (debug) LogDebugMessage($"   - 掃描到 {allButtons.Count} 個按鈕元件。");

                foreach (AutomationElement button in allButtons)
                {
                    try
                    {
                        string name = button.Current.Name;
                        if (string.IsNullOrWhiteSpace(name)) continue;

                        if (debug) LogDebugMessage($"   - 找到按鈕: '{name}'");

                        if (inCallButtonKeywords.Any(keyword => name.Contains(keyword, StringComparison.OrdinalIgnoreCase)))
                        {
                            if (debug) LogDebugMessage($"   -> ✓ 找到通話控制按鈕 '{name}'。判斷為通話視窗。");
                            return true;
                        }
                    }
                    catch (ElementNotAvailableException) { }
                }
            }
            catch (Exception ex)
            {
                if (debug) LogDebugMessage($"   - UIA 掃描期間出錯: {ex.Message}");
                return false;
            }

            if (debug) LogDebugMessage($"   -> ✗ 未找到任何通話控制按鈕。判斷為非通話視窗。");
            return false;
        }

        private void ChangeState(AppState newState, IntPtr windowHandle, List<string> teamsActions)
        {
            if (currentState == newState) return; // 避免不必要的重複狀態變更

            LogDebugMessage($"--- 狀態變更: {currentState} -> {newState} ---");
            currentState = newState;
            teamsCallWindowHandle = windowHandle;

            if (newState == AppState.InCall)
            {
                inCallStateChangeTime = DateTime.Now;
                LogDebugMessage($" -> 已進入 InCall 狀態，{IN_CALL_GRACE_PERIOD_SECONDS} 秒寬限期開始。");
            }

            this.Invoke(new Action(() => {
                UpdateButtonsForApp(teamsActions);
            }));
        }

        private string GetProcessNameFromHwnd(IntPtr hwnd)
        {
            try
            {
                GetWindowThreadProcessId(hwnd, out uint pid);
                return Process.GetProcessById((int)pid).ProcessName;
            }
            catch { return ""; }
        }

        #endregion

        #region UI 創建與事件處理

        private void UpdateButtonsForApp(List<string> teamsActions)
        {
            this.Text = $"TouchKeyBoard - {currentState}";
            switch (currentState)
            {
                case AppState.IncomingCall:
                    SetupIncomingCallButtons(teamsActions);
                    break;
                case AppState.InCall:
                    SetupInCallButtons();
                    break;
                case AppState.Normal:
                default:
                    SetupNormalButtons();
                    break;
            }
        }

        private void SetupNormalButtons()
        {
            this.Text = $"TouchKeyBoard - {currentFocusedAppName}";
            foreach (Button btn in buttons)
            {
                btn.Text = ""; btn.Tag = null; btn.Enabled = false;
                btn.BackColor = SystemColors.Control; btn.ForeColor = SystemColors.ControlText;
            }

            var appKey = appShortcuts.Keys.FirstOrDefault(k => currentFocusedAppName.ToLower().Contains(k.ToLower()));

            if (appKey == null)
            {
                if (debug) LogDebugMessage($"按鈕已清空 (無對應設定 for '{currentFocusedAppName}')");
                return;
            }

            var shortcuts = appShortcuts[appKey];
            for (int i = 0; i < shortcuts.Count && i < buttons.Length; i++)
            {
                buttons[i].Text = $"{shortcuts[i].DisplayName}\n{shortcuts[i].KeyCombination}";
                buttons[i].Tag = shortcuts[i];
                buttons[i].Enabled = true;
            }
        }

        private void SetupIncomingCallButtons(List<string> actions)
        {
            foreach (Button btn in buttons)
            {
                btn.Text = ""; btn.Tag = null; btn.Enabled = false;
                btn.BackColor = SystemColors.Control; btn.ForeColor = SystemColors.ControlText;
            }
            int buttonIndex = 0;
            if (actions.Contains("Video"))
            {
                buttons[buttonIndex].Text = "接聽視訊";
                buttons[buttonIndex].Tag = new ShortcutKey { DisplayName = "接聽視訊", KeyCombination = "Ctrl+Shift+A", SendKeysFormat = "^+a" };
                buttons[buttonIndex].Enabled = true;
                buttons[buttonIndex].BackColor = Color.FromArgb(0, 120, 212);
                buttons[buttonIndex].ForeColor = Color.White;
                buttonIndex++;
            }
            if (actions.Contains("Audio"))
            {
                buttons[buttonIndex].Text = "接聽語音";
                buttons[buttonIndex].Tag = new ShortcutKey { DisplayName = "接聽語音", KeyCombination = "Ctrl+Shift+S", SendKeysFormat = "^+s" };
                buttons[buttonIndex].Enabled = true;
                buttons[buttonIndex].BackColor = Color.FromArgb(4, 153, 114);
                buttons[buttonIndex].ForeColor = Color.White;
                buttonIndex++;
            }
            if (actions.Contains("Decline"))
            {
                buttons[buttonIndex].Text = "拒絕";
                buttons[buttonIndex].Tag = new ShortcutKey { DisplayName = "拒絕", KeyCombination = "Ctrl+Shift+D", SendKeysFormat = "^+d" };
                buttons[buttonIndex].Enabled = true;
                buttons[buttonIndex].BackColor = Color.FromArgb(217, 48, 48);
                buttons[buttonIndex].ForeColor = Color.White;
                buttonIndex++;
            }
        }

        private void SetupInCallButtons()
        {
            this.Text = "TouchKeyBoard - InCall";
            foreach (Button btn in buttons)
            {
                btn.Text = ""; btn.Tag = null; btn.Enabled = false;
                btn.BackColor = SystemColors.Control; btn.ForeColor = SystemColors.ControlText;
            }
            buttons[0].Text = "靜音切換";
            buttons[0].Tag = new ShortcutKey { DisplayName = "靜音切換", KeyCombination = "Ctrl+Shift+M", SendKeysFormat = "^+m" };
            buttons[0].Enabled = true;
            buttons[0].BackColor = Color.DarkGray;
            buttons[0].ForeColor = Color.White;
            buttons[1].Text = "視訊切換";
            buttons[1].Tag = new ShortcutKey { DisplayName = "視訊切換", KeyCombination = "Ctrl+Shift+O", SendKeysFormat = "^+o" };
            buttons[1].Enabled = true;
            buttons[1].BackColor = Color.DarkSlateBlue;
            buttons[1].ForeColor = Color.White;
            buttons[2].Text = "掛斷";
            buttons[2].Tag = new ShortcutKey { DisplayName = "掛斷", KeyCombination = "Ctrl+Shift+H", SendKeysFormat = "^+h" };
            buttons[2].Enabled = true;
            buttons[2].BackColor = Color.FromArgb(217, 48, 48);
            buttons[2].ForeColor = Color.White;
        }

        private void Button_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            if (button.Tag is not ShortcutKey shortcut) return;

            LogDebugMessage($">>> 執行快捷鍵: {shortcut.DisplayName} ({shortcut.SendKeysFormat})");

            // ▼▼▼ 核心修正：確保在 Button_Click 中使用的 handle 是正確的 ▼▼▼
            IntPtr targetWindow;
            if (currentState == AppState.IncomingCall)
            {
                targetWindow = teamsCallWindowHandle; // 來電時，目標是通知視窗
            }
            else if (currentState == AppState.InCall)
            {
                // 通話中，目標應該是前景的通話視窗，GetForegroundWindow() 最可靠
                targetWindow = GetForegroundWindow();
                LogDebugMessage($" -> InCall 點擊，使用前景視窗 {targetWindow} 作為目標");
            }
            else
            {
                targetWindow = lastActiveWindow; // 正常模式
            }
            // ▲▲▲ 核心修正 ▲▲▲

            if (targetWindow == IntPtr.Zero)
            {
                LogDebugMessage("✗ 目標窗口句柄為空，無法發送按鍵。");
                return;
            }

            try
            {
                bool switchResult = SetForegroundWindow(targetWindow);
                if (debug) LogDebugMessage($"切換窗口結果: {(switchResult ? "成功" : "失敗")}");

                if (switchResult)
                {
                    System.Threading.Thread.Sleep(100);
                    SendKeys.SendWait(shortcut.SendKeysFormat);
                    LogDebugMessage($"✓ 快捷鍵已發送");

                    if (currentState == AppState.IncomingCall)
                    {
                        if (shortcut.DisplayName.Contains("接聽"))
                        {
                            // 這裡的 targetWindow 就是我們要關閉的舊通知視窗
                            IntPtr windowToClose = targetWindow;
                            ChangeState(AppState.InCall, IntPtr.Zero, null);

                            System.Windows.Forms.Timer closeTimer = new System.Windows.Forms.Timer();
                            closeTimer.Interval = 1500;
                            closeTimer.Tick += (timerSender, timerArgs) =>
                            {
                                SendMessage(windowToClose, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                                LogDebugMessage($"✓ [延遲任務] 已發送 WM_CLOSE 訊息到來電視窗 {windowToClose}。");
                                closeTimer.Stop();
                                closeTimer.Dispose();
                            };
                            closeTimer.Start();
                            LogDebugMessage($"-> 已啟動 1.5 秒延遲關閉計時器。");
                        }
                        else if (shortcut.DisplayName.Contains("拒絕"))
                        {
                            ChangeState(AppState.Normal, IntPtr.Zero, null);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (debug) LogDebugMessage($"✗ 發送快捷鍵失敗: {ex.Message}");
            }
        }

        private void CreateButtons()
        {
            int buttonHeight = ClientSize.Height / 10;

            for (int i = 0; i < 10; i++)
            {
                buttons[i] = new Button
                {
                    Width = ClientSize.Width,
                    Height = buttonHeight,
                    Location = new Point(0, i * buttonHeight),
                    Text = "",
                    Enabled = false,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
                    TabStop = false,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 9F, FontStyle.Bold)
                };
                buttons[i].FlatAppearance.BorderSize = 1;
                buttons[i].Click += Button_Click;
                Controls.Add(buttons[i]);
            }

            Button configButton = new Button
            {
                Size = new Size(80, 30),
                Location = new Point(ClientSize.Width - 85, 5),
                Text = "設定",
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                TabStop = false
            };
            configButton.Click += ConfigButton_Click;
            Controls.Add(configButton);

            Resize += Form1_Resize;
        }

        private void ConfigButton_Click(object sender, EventArgs e)
        {
            if (debug) LogDebugMessage("設定按鈕被點擊");
            MessageBox.Show($"您可以編輯設定檔來自訂快捷鍵：\n{configFilePath}", "設定");
            try
            {
                Process.Start("explorer.exe", $"/select,\"{configFilePath}\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"無法打開檔案總管：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            int buttonHeight = ClientSize.Height / 10;
            for (int i = 0; i < 10; i++)
            {
                if (buttons[i] != null)
                {
                    buttons[i].Width = ClientSize.Width;
                    buttons[i].Height = buttonHeight;
                    buttons[i].Location = new Point(0, i * buttonHeight);
                }
            }
        }

        #endregion
    }

    public class ShortcutKey
    {
        public string DisplayName { get; set; } = "";
        public string KeyCombination { get; set; } = "";
        public string SendKeysFormat { get; set; } = "";
        public string Description { get; set; } = "";
    }
}