// TeamsDetector.cs
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Automation;

namespace TouchKeyBoard
{
    #region 焦點管理器 (FocusManager)
    /// <summary>
    /// 一個靜態工具類別，提供主動檢查前景視窗資訊的功能。
    /// 主要用於判斷當前焦點所在的視窗是否為 Teams 的通話中視窗，
    /// 或從視窗控制代碼 (HWND) 取得其對應的行程名稱。
    /// </summary>
    public static class FocusManager
    {
        #region 私有常數

        /// <summary>
        /// 定義 Microsoft Teams 的行程名稱，用於比對。
        /// </summary>
        private const string TEAMS_PROCESS_NAME = "ms-teams";

        #endregion

        #region 公開靜態方法

        /// <summary>
        /// 根據提供的視窗控制代碼 (HWND) 獲取其所屬的行程名稱。
        /// </summary>
        /// <param name="hwnd">目標視窗的控制代碼。</param>
        /// <returns>行程的名稱 (例如 "ms-teams")，如果失敗則返回空字串。</returns>
        public static string GetProcessNameFromHwnd(IntPtr hwnd)
        {
            try
            {
                // 透過 Win32 API 取得視窗的執行緒與行程 ID (PID)
                NativeMethods.GetWindowThreadProcessId(hwnd, out uint pid);
                // 根據 PID 取得完整的 Process 物件，並返回其名稱
                return Process.GetProcessById((int)pid).ProcessName;
            }
            catch
            {
                // 發生任何錯誤 (例如行程已關閉) 則返回空字串
                return "";
            }
        }

        /// <summary>
        /// 檢查指定的視窗是否為 Teams 的「通話中」主視窗。
        /// 此方法透過 UI Automation 掃描視窗內的按鈕元件，
        /// 如果找到符合通話中特徵 (如 "靜音"、"掛斷") 的按鈕，則判定為是。
        /// </summary>
        /// <param name="hwnd">要檢查的視窗控制代碼。</param>
        /// <returns>如果視窗是 Teams 通話中視窗，返回 true；否則返回 false。</returns>
        public static bool IsTeamsInCallWindow(IntPtr hwnd)
        {
            // 步驟 1: 檢查視窗是否屬於 Teams 行程
            string processName = GetProcessNameFromHwnd(hwnd);
            if (!processName.Equals(TEAMS_PROCESS_NAME, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // 步驟 2: 定義通話中視窗會出現的按鈕關鍵字 (多國語言)
            string[] inCallButtonKeywords = {
                "Mute", "Unmute", "靜音",
                "Camera", "攝影機", "Turn camera on", "Turn camera off",
                "Leave", "Hang up", "離開", "掛斷",
                "Share", "共用"
            };

            try
            {
                // 步驟 3: 使用 UI Automation 技術來分析視窗內容
                AutomationElement rootElement = AutomationElement.FromHandle(hwnd);
                if (rootElement == null) return false;

                // 建立一個查詢條件，只尋找類型為「按鈕」的元件
                var buttonCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                // 尋找視窗內所有後代按鈕元件
                var allButtons = rootElement.FindAll(TreeScope.Descendants, buttonCondition);

                // 步驟 4: 遍歷所有找到的按鈕
                foreach (AutomationElement button in allButtons)
                {
                    try
                    {
                        // 取得按鈕的名稱 (通常是按鈕上的文字或無障礙標籤)
                        string name = button.Current.Name;
                        if (string.IsNullOrWhiteSpace(name)) continue;

                        // 比對按鈕名稱是否包含任何一個通話中的關鍵字
                        if (inCallButtonKeywords.Any(keyword => name.Contains(keyword, StringComparison.OrdinalIgnoreCase)))
                        {
                            // 只要找到一個符合的，就立刻返回 true
                            return true;
                        }
                    }
                    catch (ElementNotAvailableException)
                    {
                        // 如果在檢查過程中，該元件突然消失，則忽略此錯誤並繼續
                    }
                }
            }
            catch (Exception)
            {
                // 捕捉所有 UI Automation 可能發生的其他例外，並安全地返回 false
                return false;
            }

            // 如果掃描完所有按鈕都沒找到關鍵字，則返回 false
            return false;
        }

        #endregion
    }
    #endregion

    #region Teams 來電偵測器 (TeamsDetector)
    /// <summary>
    /// 一個負責在背景監聽系統事件的類別，專門用於偵測 Microsoft Teams 的來電通知。
    /// 它使用 WinEventHook 來被動監聽視窗的「顯示」與「隱藏」事件，
    /// 當偵測到符合條件的 Teams 來電視窗時，會觸發對應的公開事件。
    /// 這個類別需要被 Dispose 以釋放系統掛鉤資源。
    /// </summary>
    public class TeamsDetector : IDisposable
    {
        #region 私有常數與欄位

        /// <summary>
        /// 定義 Teams 來電通知視窗特有的類別名稱 (ClassName)，這是事件過濾的關鍵。
        /// </summary>
        private const string TEAMS_CALL_WINDOW_CLASS = "TeamsWebView";

        /// <summary>
        /// 用於儲存 WinEventProc 回呼方法的委派，防止被記憶體回收機制回收。
        /// </summary>
        private NativeMethods.WinEventDelegate _teamsCallDelegate;

        /// <summary>
        /// 儲存由 SetWinEventHook API 返回的事件掛鉤控制代碼，用於之後的卸載。
        /// </summary>
        private IntPtr _teamsEventHook = IntPtr.Zero;

        /// <summary>
        /// 一個委派，用於接收外部傳入的記錄訊息方法。
        /// </summary>
        private Action<string> _logAction;

        #endregion

        #region 公開事件

        /// <summary>
        /// 當偵測到 Teams 來電通知視窗出現時觸發。
        /// 提供視窗的控制代碼 (IntPtr) 和透過 UI Automation 分析出的可用操作列表 (List<string>)。
        /// </summary>
        public event Action<IntPtr, List<string>> CallDetected;

        /// <summary>
        /// 當先前偵測到的 Teams 來電通知視窗被隱藏或關閉時觸發。
        /// 提供視窗的控制代碼 (IntPtr)。
        /// </summary>
        public event Action<IntPtr> CallEnded;

        #endregion

        #region 建構子與啟動

        /// <summary>
        /// 初始化 TeamsDetector 的新執行個體。
        /// </summary>
        /// <param name="logAction">可選的記錄方法委派，用於輸出偵錯或狀態訊息。</param>
        public TeamsDetector(Action<string> logAction = null)
        {
            // 如果未提供記錄方法，則使用一個不做任何事的空方法，避免後續呼叫時產生 null 參考錯誤
            _logAction = logAction ?? ((_) => { });
        }

        /// <summary>
        /// 開始監聽系統事件。此方法會設定一個全域的 WinEventHook。
        /// </summary>
        public void Start()
        {
            _teamsCallDelegate = new NativeMethods.WinEventDelegate(WinEventProc);
            _teamsEventHook = NativeMethods.SetWinEventHook(
                NativeMethods.EVENT_OBJECT_SHOW,      // 監聽的最小事件 ID (視窗顯示)
                NativeMethods.EVENT_OBJECT_HIDE,      // 監聽的最大事件 ID (視窗隱藏)
                IntPtr.Zero,                          // 使用本機的 DLL
                _teamsCallDelegate,                   // 事件發生時的回呼方法
                0, 0,                                 // 監聽所有行程與所有執行緒
                NativeMethods.WINEVENT_OUTOFCONTEXT); // 非同步事件處理
            _logAction("Teams 來電監聽已啟動。");
        }

        #endregion

        #region 核心事件處理程序

        /// <summary>
        /// WinEventHook 的回呼方法。當系統發生視窗顯示或隱藏事件時，此方法會被 Windows 呼叫。
        /// </summary>
        private void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            // 過濾掉不必要的事件 (例如子物件的變化)，只關心視窗本身 (hwnd)
            if (idObject != 0 || hwnd == IntPtr.Zero) return;

            // 取得視窗的類別名稱
            StringBuilder className = new StringBuilder(256);
            NativeMethods.GetClassName(hwnd, className, className.Capacity);

            // 檢查類別名稱是否為我們鎖定的 Teams 來電視窗
            if (className.ToString() == TEAMS_CALL_WINDOW_CLASS)
            {
                if (eventType == NativeMethods.EVENT_OBJECT_SHOW)
                {
                    // 如果是「顯示」事件，則進一步分析視窗內容
                    var availableActions = DetectTeamsCallActions(hwnd);
                    if (availableActions.Count > 0)
                    {
                        // 如果分析出可用的操作 (接聽/拒絕)，則觸發 CallDetected 事件
                        CallDetected?.Invoke(hwnd, availableActions);
                    }
                }
                else if (eventType == NativeMethods.EVENT_OBJECT_HIDE)
                {
                    // 如果是「隱藏」事件，則觸發 CallEnded 事件
                    CallEnded?.Invoke(hwnd);
                }
            }
        }

        #endregion

        #region UI 自動化分析






        /// <summary>
        /// 使用 UI Automation 分析指定的視窗，偵測來電通知中的可用操作按鈕。
        /// </summary>
        /// <param name="hwnd">來電通知視窗的控制代碼。</param>
        /// <returns>一個包含 "Video", "Audio", "Decline" 的字串列表。</returns>
        private List<string> DetectTeamsCallActions(IntPtr hwnd)
        {
            var availableActions = new List<string>();

            // 初步過濾：如果視窗尺寸太小，不太可能是來電通知，直接返回
            if (!NativeMethods.GetWindowRect(hwnd, out var rect) || (rect.Right - rect.Left) <= 300) return availableActions;

            _logAction($" -> 尺寸符合. 開始 UIA 掃描來電操作...");
            try
            {
                // 與 FocusManager 類似，使用 UI Automation 掃描按鈕
                AutomationElement rootElement = AutomationElement.FromHandle(hwnd);
                if (rootElement == null) return availableActions;

                var buttonCondition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                var allButtons = rootElement.FindAll(TreeScope.Descendants, buttonCondition);

                // 定義來電操作的按鈕關鍵字
                string[] videoAcceptNames = { "Accept with video", "接聽視訊" };
                string[] audioAcceptNames = { "Accept with audio", "接聽語音", "Accept", "接聽" };
                string[] declineNames = { "Decline", "拒絕" };

                _logAction($"   - 掃描到 {allButtons.Count} 個按鈕元件。");
                foreach (AutomationElement button in allButtons)
                {
                    try
                    {
                        string name = button.Current.Name;
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        _logAction($"   - 找到按鈕: '{name}'");

                        // 根據按鈕名稱將偵測到的操作加入列表中，並確保不重複
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
                _logAction($"   - UIA 掃描出錯: {ex.Message}");
            }
            _logAction($" -> 掃描完成. 偵測到的操作: [{string.Join(", ", availableActions)}]");
            return availableActions;
        }

        #endregion

        #region 資源釋放

        /// <summary>
        /// 釋放由這個類別管理的非受控資源，主要是卸載 WinEventHook。
        /// </summary>
        public void Dispose()
        {
            if (_teamsEventHook != IntPtr.Zero)
            {
                NativeMethods.UnhookWinEvent(_teamsEventHook);
                _logAction("Teams 來電監聽掛鉤已卸載。");
                _teamsEventHook = IntPtr.Zero; // 避免重複卸載
            }
        }

        #endregion
    }
    #endregion

    public static class WindowInspector
    {
        /// <summary>
        /// 獲取指定視窗控制代碼的詳細除錯資訊。
        /// </summary>
        /// <param name="hWnd">目標視窗的控制代碼。</param>
        /// <returns>一個格式化的字串，包含視窗標題、類別、行程等資訊。</returns>
        public static string GetWindowDebugInfo(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero)
            {
                return "--- 目標視窗 Handle 為空 (IntPtr.Zero) ---";
            }

            // 獲取視窗標題
            StringBuilder titleBuilder = new StringBuilder(256);
            NativeMethods.GetWindowText(hWnd, titleBuilder, titleBuilder.Capacity);
            string title = string.IsNullOrWhiteSpace(titleBuilder.ToString()) ? "(無標題)" : titleBuilder.ToString();

            // 獲取類別名稱
            StringBuilder classBuilder = new StringBuilder(256);
            NativeMethods.GetClassName(hWnd, classBuilder, classBuilder.Capacity);
            string className = classBuilder.ToString();

            // 獲取行程資訊
            string processInfo = "(無法取得行程)";
            try
            {
                NativeMethods.GetWindowThreadProcessId(hWnd, out uint pid);
                Process p = Process.GetProcessById((int)pid);
                processInfo = $"{p.ProcessName}.exe (PID: {pid})";
            }
            catch { /* 忽略錯誤 */ }

            // 格式化輸出
            var sb = new StringBuilder();
            sb.AppendLine("▼▼▼ 目標視窗資訊 ▼▼▼");
            sb.AppendLine($"  Handle: {hWnd}");
            sb.AppendLine($"  標題 (Title): {title}");
            sb.AppendLine($"  類別 (Class): {className}");
            sb.AppendLine($"  所屬行程: {processInfo}");
            sb.AppendLine("▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲");

            return sb.ToString();
        }
    }
}