using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    public class WebViewInterop {
        public void jsOpenFile(string fileName) {
            App.OpenFile(fileName);
        }
    }
}
