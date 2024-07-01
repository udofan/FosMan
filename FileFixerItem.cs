using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    internal class FileFixerItem {
        public string DirectoryName { get; set; }
        public string FileName { get; set; }
        public string FullFileName { get; set; }
        
        public FileFixerItem(string file) {
            FullFileName = file;
            DirectoryName = Path.GetDirectoryName(file);
            FileName = Path.GetFileName(file);
        }
    }
}
