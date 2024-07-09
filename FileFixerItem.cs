using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static FosMan.Enums;

namespace FosMan {
    internal class FileFixerItem {
        public string DirectoryName { get; set; }
        public string FileName { get; set; }
        public string FullFileName { get; set; }
        public EFileType Type { get; set; }
        public Rpd Rpd { get; set; }
        public Fos Fos { get; set; }

        public FileFixerItem(string file, EFileType fileType) {
            FullFileName = file;
            DirectoryName = Path.GetDirectoryName(file);
            FileName = Path.GetFileName(file);

            if (fileType == EFileType.Auto) {
                if (FileName.Contains("РПД")) {
                    fileType = EFileType.Rpd;
                }
                else if (FileName.Contains("ФОС")) {
                    fileType = EFileType.Fos;
                }
                else {
                    fileType = EFileType.Unknown;
                }
            }

            Type = fileType;
        }

        public FileFixerItem(Rpd rpd) {
            Rpd = rpd;
            Type = EFileType.Rpd;
            FullFileName = rpd.SourceFileName;
            DirectoryName = Path.GetDirectoryName(FullFileName);
            FileName = Path.GetFileName(FullFileName);
        }
        public FileFixerItem(Fos fos) {
            Fos = fos;
            Type = EFileType.Fos;
            FullFileName = fos.SourceFileName;
            DirectoryName = Path.GetDirectoryName(FullFileName);
            FileName = Path.GetFileName(FullFileName);
        }
    }
}
