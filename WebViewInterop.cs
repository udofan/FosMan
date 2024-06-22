using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    public class WebViewInterop {
        public void jsOpenFile(string fileName) {
            try {
                var p = new Process();
                p.StartInfo = new ProcessStartInfo(fileName) {
                    UseShellExecute = true
                };
                p.Start();
            }
            catch {
                MessageBox.Show($"Не удалось открыть файл\n{fileName}", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
