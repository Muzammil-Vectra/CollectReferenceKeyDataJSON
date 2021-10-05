using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CollectReferenceKeyDataJSON
{
    class TopologyMap
    {
        public TopologyMap()
        {
            BodyIdentifiers = new Dictionary<string, string>();
            FaceIdentifiers = new Dictionary<string, string>();
            EdgeIdentifiers = new Dictionary<string, string>();
        }
        public string PartNumber { get; set; }
        public string PartIdentifier { get; set; }
        public string AssemblyNumber { get; set; }
        public string AssemblyIdentifier { get; set; }
        public Dictionary<string, string> BodyIdentifiers { get; set; }
        public Dictionary<string, string> FaceIdentifiers { get; set; }
        public Dictionary<string, string> EdgeIdentifiers { get; set; }
    }
}
