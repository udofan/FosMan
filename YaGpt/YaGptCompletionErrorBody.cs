using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Pkcs;
using System.Text;
using System.Threading.Tasks;

namespace FosMan
{
    internal class YaGptCompletionErrorBody
    {
        //{"error": {
        //  "grpcCode":8,
        //  "httpCode":429,
        //  "message":"ai.textGenerationCompletionSessionsCount.count gauge quota limit exceed: allowed 1 requests",
        //  "httpStatus":"Too Many Requests",
        //  "details":[]
        //  }
        //}
        public int grpcCode { get; set; }
        public int httpCode { get; set; }
        public string message { get; set; }
        public string httpStatus { get; set; }
        public string[] details { get; set; }
    }
}
