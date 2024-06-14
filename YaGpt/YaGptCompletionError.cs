using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FosMan
{
    internal class YaGptCompletionError
    {
        //{"error": {
        //  "grpcCode":8,
        //  "httpCode":429,
        //  "message":"ai.textGenerationCompletionSessionsCount.count gauge quota limit exceed: allowed 1 requests",
        //  "httpStatus":"Too Many Requests",
        //  "details":[]
        //  }
        //}
        public YaGptCompletionErrorBody error { get; set; }
    }
}
