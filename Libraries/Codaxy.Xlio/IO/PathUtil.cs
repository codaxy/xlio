using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.IO
{
    class PathUtil
    {
        public static void SplitFilePath(string path, out String filePath, out String fileName)
        {
            var ind = path.LastIndexOf('/');
            if (ind == -1)
            {
                filePath = "";
                fileName = path;
            }
            else
            {
                filePath = path.Substring(0, ind);
                fileName = path.Substring(ind + 1);
            }
        }

        public static String CombinePaths(string path1, String path2)
        {
            return path1.TrimEnd('/') + "/" + path2.TrimStart('/');
        }
    }
}
