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

        public static bool AreEqual(object o1, object o2)
        {
            if (o1 == null)
                return o2 == null;
            return System.Object.ReferenceEquals(o1, o2) || o1.Equals(o2);
        }
    }
}
