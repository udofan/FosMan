using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan {
    static internal class CurriculumReader {
        public static List<CurriculumProperty> @Properties { get; } = [
            new CurriculumProperty() {
                HeaderText = "ИНДЕКС",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "Index"
            },
            new CurriculumProperty() {
                HeaderText = "НАИМЕНОВАНИЕ",
                TestFunction = EPropertyTestFunction.Contains,
                TargetProperty = "Name"
            }
        ];
    }
}
