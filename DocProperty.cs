using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    internal class DocProperty {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
        public bool IsChecked { get; set; }

        public static List<DocProperty> DefaultProperties { get; } = new() {
            new() { Name = "dc:title", Value = "" },
            new() { Name = "dc:creator", Value = "Красавина В.А." },
            new() { Name = "cp:lastModifiedBy", Value = "Красавина В.А." },
            new() { Name = "cp:lastPrinted", Value = "" },
            new() { Name = "dcterms:created", Value = "2024-05-20T08:46:00Z" },
            new() { Name = "dcterms:modified", Value = "2024-05-20T10:49:00Z" }
        };
    }
}
