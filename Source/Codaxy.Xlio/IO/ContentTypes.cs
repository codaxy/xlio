using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.IO
{
    //SRP: Runtime ContentType resolution
    public partial class ContentTypes
    {
        Dictionary<String, String> defaultType;
        Dictionary<String, String> partType;

        public ContentTypes()
        {
            defaultType = new Dictionary<string, string>();
            partType = new Dictionary<string, string>();
        }

        public void RegisterDefaultType(String ext, String type)
        {
            defaultType[ext] = type;
        }

        public void RegisterPartTypeOverride(String partName, String type)
        {
            var path = partName.TrimStart('/');
            partType[path] = type;
        }

        public String ResolvePartType(String partName)
        {
            String res;
            if (partType.TryGetValue(partName, out res))
                return res;
            var ext = System.IO.Path.GetExtension(partName).TrimStart('.');
            if (defaultType.TryGetValue(ext, out res))
                return res;
            return null;
        }

    }
}
