using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Codaxy.Xlio.IO
{
    //Copied from Common.Reflection.TypeInfo
    public class TypeInfo
    {
        public static bool IsNullableType(Type type)
        {
            return (type.IsGenericType && type.GetGenericTypeDefinition().Equals(typeof(Nullable<>)));
        }

        static Type[] numericTypes = new[] { typeof(int), typeof(decimal), typeof(float), typeof(uint), typeof(long), typeof(ulong), typeof(double), typeof(short), typeof(ushort) };

        public static bool IsNumericType(Type type)
        {
            if (IsNullableType(type))
                return IsNumericType(Nullable.GetUnderlyingType(type));
            return numericTypes.Contains(type);
        }
    }
}
