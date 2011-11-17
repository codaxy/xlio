using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    class Util
    {
        public static bool TryCast<T>(object o, out T t)
        {
            if (o is T)
            {
                t = (T)o;
                return true;
            }
            t = default(T);
            return false;
        }
    }
}
