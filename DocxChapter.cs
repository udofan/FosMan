using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace FosMan {
    internal class DocxChapter {
        public string Number { get; set; }
        public string Text { get; set; }
        public List<Paragraph> Paragraphs { get; set; }
    }
}
