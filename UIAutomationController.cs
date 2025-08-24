// UIAutomationController.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Automation;

namespace TouchKeyBoard
{
    public static class UIAutomationController
    {
        /// <summary>
        /// 在指定的父視窗內尋找名稱匹配的按鈕，並以程式化方式「點擊」它。
        /// </summary>
        /// <param name="parentHwnd">目標按鈕所在的父視窗控制代碼。</param>
        /// <param name="targetNames">一個包含多個可能按鈕名稱的列表（用於應對多語言）。</param>
        /// <param name="logAction">用於記錄除錯訊息的委派。</param>
        /// <returns>如果成功找到並點擊了按鈕，返回 true；否則返回 false。</returns>
        public static bool InvokeButtonByName(IntPtr parentHwnd, List<string> targetNames, Action<string> logAction)
        {
            if (parentHwnd == IntPtr.Zero)
            {
                logAction?.Invoke("✗ UIA 操作失敗：父視窗 Handle 為空。");
                return false;
            }

            try
            {
                // 1. 從視窗控制代碼獲取 UI Automation 的根元素
                AutomationElement rootElement = AutomationElement.FromHandle(parentHwnd);
                if (rootElement == null)
                {
                    logAction?.Invoke($"✗ UIA 操作失敗：無法從 Handle {parentHwnd} 獲取 AutomationElement。");
                    return false;
                }

                // 2. 設定搜尋條件：我們只對「按鈕」類型的控制項感興趣
                Condition condition = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);

                // 3. 尋找視窗內所有的後代按鈕
                AutomationElementCollection buttons = rootElement.FindAll(TreeScope.Descendants, condition);
                logAction?.Invoke($"-> UIA 掃描：在視窗 {parentHwnd} 中找到 {buttons.Count} 個按鈕元件。");

                // 4. 遍歷所有找到的按鈕，尋找名稱匹配的目標
                foreach (AutomationElement button in buttons)
                {
                    string name = button.Current.Name;
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    // 檢查按鈕名稱是否包含我們目標列表中的任何一個關鍵字
                    if (targetNames.Any(target => name.Contains(target, StringComparison.OrdinalIgnoreCase)))
                    {
                        logAction?.Invoke($"  ✓ 找到目標按鈕: '{name}'");

                        // 5. 【核心】獲取按鈕的「Invoke」模式
                        if (button.TryGetCurrentPattern(InvokePattern.Pattern, out object pattern))
                        {
                            InvokePattern invokePattern = (InvokePattern)pattern;

                            // 6. 【核心】執行 Invoke，這等同於一次程式化的點擊
                            invokePattern.Invoke();
                            logAction?.Invoke($"  ✓✓✓ 成功透過 UIA Invoke 點擊了按鈕: '{name}'");
                            return true; // 成功後立刻返回
                        }
                        else
                        {
                            logAction?.Invoke($"  ✗ 錯誤：找到按鈕 '{name}' 但它不支援 InvokePattern。");
                        }
                    }
                }

                logAction?.Invoke("✗ UIA 掃描完成，但未能找到任何名稱匹配的目標按鈕。");
                return false;
            }
            catch (Exception ex)
            {
                logAction?.Invoke($"✗ UIA 操作期間發生嚴重錯誤: {ex.Message}");
                return false;
            }
        }
    }
}