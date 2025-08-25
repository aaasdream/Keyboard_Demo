// ConfigManager.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows.Forms;

namespace TouchKeyBoard
{
    public class ConfigManager
    {
        private readonly string _configFilePath;
        private readonly Action<string> _logAction;

        public ConfigManager(string configFilePath, Action<string> logAction = null)
        {
            _configFilePath = configFilePath;
            _logAction = logAction ?? ((_) => { });
        }

        public Dictionary<string, List<ShortcutKey>> LoadShortcuts()
        {
            try
            {
                if (File.Exists(_configFilePath))
                {
                    _logAction("找到設定檔，正在讀取...");
                    string json = File.ReadAllText(_configFilePath);
                    var loadedSettings = JsonSerializer.Deserialize<Dictionary<string, List<ShortcutKey>>>(json);

                    if (loadedSettings != null)
                    {
                        _logAction($"✓ 設定檔載入成功，共 {loadedSettings.Count} 個應用程式設定。");
                        return loadedSettings;
                    }

                    _logAction("✗ 設定檔為空或格式錯誤，將使用空設定。");
                    return new Dictionary<string, List<ShortcutKey>>();
                }

                _logAction("設定檔不存在，將創建並載入預設設定檔。");
                CreateDefaultConfigFile();
                return LoadShortcuts(); // 重新載入
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加載設定檔時出錯: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _logAction($"✗ 載入設定檔失敗: {ex.Message}");
                return new Dictionary<string, List<ShortcutKey>>();
            }
        }

        private void CreateDefaultConfigFile()
        {
            try
            {
                string defaultConfigJson = GetDefaultConfigJsonString();
                File.WriteAllText(_configFilePath, defaultConfigJson);
                _logAction("✓ 預設設定檔已創建成功。");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"創建預設設定檔時出錯: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _logAction($"✗ 創建預設設定檔失敗: {ex.Message}");
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
              ],
              "ms-teams": [
              { "DisplayName": "搜尋", "KeyCombination": "Ctrl+E", "SendKeysFormat": "^e", "Description": "跳至搜尋框" },
              { "DisplayName": "新增聊天", "KeyCombination": "Ctrl+N", "SendKeysFormat": "^n", "Description": "開始新的聊天" },
              { "DisplayName": "開啟設定", "KeyCombination": "Ctrl+,", "SendKeysFormat": "^,", "Description": "開啟設定" },
              { "DisplayName": "開啟說明", "KeyCombination": "F1", "SendKeysFormat": "{F1}", "Description": "開啟說明" },
              { "DisplayName": "放大", "KeyCombination": "Ctrl+=", "SendKeysFormat": "^=", "Description": "放大畫面" },
              { "DisplayName": "縮小", "KeyCombination": "Ctrl+-", "SendKeysFormat": "^-", "Description": "縮小畫面" },
              { "DisplayName": "附加檔案", "KeyCombination": "Ctrl+O", "SendKeysFormat": "^o", "Description": "附加檔案" },
              { "DisplayName": "切換至聊天", "KeyCombination": "Ctrl+2", "SendKeysFormat": "^2", "Description": "" },
              { "DisplayName": "切換至行事曆", "KeyCombination": "Ctrl+4", "SendKeysFormat": "^4", "Description": "" },
              { "DisplayName": "切換至通話", "KeyCombination": "Ctrl+5", "SendKeysFormat": "^5", "Description": "" }
              ]
            }
            """;
        }
    }
}