using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    class HashCodeHelper
    {
        public static int CalculateHashCode(params object[] objects)
        {
            unchecked
            {
                int res = 17;
                foreach (var o in objects)
                {
                    res *= 23;
                    if (o != null)
                        res += o.GetHashCode();
                }
                return res;
            }
        }
    }
}
