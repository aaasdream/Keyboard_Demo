using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// ShortcutKey.cs
namespace TouchKeyBoard
{
    public class ShortcutKey
    {
        public string DisplayName { get; set; } = "";
        public string KeyCombination { get; set; } = "";
        public string SendKeysFormat { get; set; } = "";
        public string Description { get; set; } = "";
    }
}