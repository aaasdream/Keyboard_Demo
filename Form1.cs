using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace TouchKeyBoard
{
    public partial class Form1 : Form
    {
        #region Win32 API 導入與常數
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;


        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_NOACTIVATE = 0x08000000;

        #endregion

        #region 變數與屬性

        private bool debug = true;
        private Form debugWindow;
        private TextBox debugTextBox;
        private Button[] buttons = new Button[10];
        private string currentFocusedAppName = "";
        private IntPtr lastActiveWindow = IntPtr.Zero;
        private readonly string configFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "config.json"
        );
        private Dictionary<string, List<ShortcutKey>> appShortcuts = new();
        private System.Windows.Forms.Timer focusCheckTimer = new();

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
            }
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
            
            this.FormClosed += (s, e) => debugWindow?.Close();
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

        /// <summary>
        /// 載入設定檔。如果檔案不存在，則創建一個預設的設定檔，然後再載入。
        /// </summary>
        private void LoadConfig()
        {
            try
            {
                // --- 步驟 1: 檢查設定檔是否存在 ---
                if (File.Exists(configFilePath))
                {
                    // 檔案存在，直接讀取
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
                    // --- 步驟 2: 檔案不存在，進入創建流程 ---
                    if (debug) LogDebugMessage("設定檔不存在，將創建並載入預設設定檔。");
                    
                    // 2a. 呼叫方法，將預設的 JSON 字串寫入新檔案
                    CreateDefaultConfigFile();
                    
                    // 2b. 重新呼叫自己，這次就會讀取到剛剛創建的檔案
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

        /// <summary>
        /// 當 config.json 不存在時，創建一個包含預設內容的檔案。
        /// </summary>
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

        /// <summary>
        /// 返回包含預設快捷鍵的 JSON 格式字串。
        /// </summary>
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

        #region 焦點監控與邏輯

        private void CheckFocusedWindow(object sender, EventArgs e)
        {
            //
            SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);


            IntPtr currentWindow = GetForegroundWindow();
            if (currentWindow == this.Handle || (debugWindow != null && currentWindow == debugWindow.Handle))
            {
                return;
            }

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
                    
                    this.Text = $"TouchKeyBoard - {appName}";
                    
                    if (debug)
                    {
                        LogDebugMessage($"--- 視窗焦點變更: {appName} ---");
                        if (appShortcuts.ContainsKey(appName))
                        {
                            LogDebugMessage($"★★★ 找到匹配設定: {appName}，共 {appShortcuts[appName].Count} 個快捷鍵 ★★★");
                        }
                        else
                        {
                            LogDebugMessage($"   (無匹配設定)");
                        }
                    }
                    
                    UpdateButtonsForApp(appName);
                }
            }
        }

        private string GetProcessNameFromHwnd(IntPtr hwnd)
        {
            try
            {
                GetWindowThreadProcessId(hwnd, out uint pid);
                Process process = Process.GetProcessById((int)pid);
                return process.ProcessName;
            }
            catch
            {
                return "";
            }
        }
        
        #endregion

        #region UI 創建與事件處理

        private void UpdateButtonsForApp(string appName)
        {
            foreach (Button btn in buttons)
            {
                if (btn != null)
                {
                    btn.Text = "";
                    btn.Tag = null;
                    btn.Enabled = false;
                }
            }
            
            if (!appShortcuts.ContainsKey(appName))
            {
                if (debug) LogDebugMessage("按鈕已清空（無對應設定）");
                return;
            }
                
            var shortcuts = appShortcuts[appName];
            int enabledButtons = 0;
            for (int i = 0; i < shortcuts.Count && i < buttons.Length; i++)
            {
                buttons[i].Text = $"{shortcuts[i].DisplayName}\n{shortcuts[i].KeyCombination}";
                buttons[i].Tag = shortcuts[i];
                buttons[i].Enabled = true;
                enabledButtons++;
            }
            
            if (debug) LogDebugMessage($"✓ 按鈕更新完成，啟用 {enabledButtons} 個按鈕");
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
                    FlatStyle = FlatStyle.Flat
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

        private void Button_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            if (button.Tag is not ShortcutKey shortcut) return;
            
            if (debug)
            {
                LogDebugMessage($">>> 執行快捷鍵: {shortcut.DisplayName} ({shortcut.KeyCombination})");
                LogDebugMessage($"目標窗口句柄: {lastActiveWindow}");
            }
            
            if (lastActiveWindow != IntPtr.Zero)
            {
                try
                {
                    bool switchResult = SetForegroundWindow(lastActiveWindow);
                    if (debug) LogDebugMessage($"切換窗口結果: {(switchResult ? "成功" : "失敗")}");
                    
                    if (switchResult)
                    {
                        System.Threading.Thread.Sleep(100); 
                        SendKeys.SendWait(shortcut.SendKeysFormat);
                        if (debug) LogDebugMessage($"✓ 快捷鍵已發送到目標窗口: {shortcut.SendKeysFormat}");
                    }
                }
                catch (Exception ex)
                {
                    if (debug) LogDebugMessage($"✗ 發送快捷鍵失敗: {ex.Message}");
                }
            }
            else
            {
                if (debug) LogDebugMessage("✗ 沒有記住的目標窗口");
            }
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