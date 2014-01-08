using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio
{
    class ExceptionFactory
    {
        public static ArgumentException Argument(String format, params object[] args)
        {
            return new ArgumentException(String.Format(format, args));
        }

        public static InvalidOperationException InvalidOperation(String format, params object[] args)
        {
            return new InvalidOperationException(String.Format(format, args));
        }       
    }
}
