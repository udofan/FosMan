using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace FosMan {
    internal class Config {
        [JsonInclude]
        public string CompetenceMatrixFileName { get; set; }
        [JsonInclude]
        public string RpdLastLocation { get; set; }
        [JsonInclude]
        public string CurriculumLastLocation { get; set; }
        [JsonInclude]
        public List<RpdFindAndReplaceItem> RpdFindAndReplaceItems { get; set; } = [];
        [JsonInclude]
        public List<CurriculumDisciplineHeader> CurriculumDisciplineParseItems { get; set; }
    }
}
