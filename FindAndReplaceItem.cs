using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    internal class FindAndReplaceItem {
        public bool IsChecked { get; set; }
        public string FindPattern { get; set; } = "";
        public string ReplacePattern { get; set; } = "";
    }
}
