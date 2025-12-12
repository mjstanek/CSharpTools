using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataverseMappingRemover.Services
{
    public class MappingResult
    {
        public string SourceEntity { get; set; } = "";
        public string SourceAttribute { get; set; } = "";
        public string TargetEntity { get; set; } = "";
        public string TargetAttribute { get; set; } = "";
        public Guid AttributeMapId { get; set; } = Guid.Empty;
        public Guid EntityMapId { get; set; } = Guid.Empty;
        public string Direction { get; set; } = "Source --> Target";
    }
}
