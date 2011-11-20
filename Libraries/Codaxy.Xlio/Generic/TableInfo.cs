using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Codaxy.Xlio;
using System.Globalization;
using Codaxy.Xlio.IO;

namespace Codaxy.Xlio.Generic
{
    public class TableInfo<T>
    {
        public List<ColumnInfo<T>> Columns { get; set; }        

        public static TableInfo<T> Build()
        {
            var type = typeof(T);
            var properties = type.GetProperties();
            var result = new TableInfo<T>
            {
                Columns = new List<ColumnInfo<T>>()
            };

            foreach (var p in properties)
            {
                var prop = p;
                result.Columns.Add(new ColumnInfo<T>
                {
                    Name = prop.Name,
                    Getter = prop.CanRead ? (row) => { return prop.GetValue(row, null); } : (Func<T, object>)null,
                    Setter = prop.CanWrite ? (row, value) =>
                    {                        
                        prop.SetValue(row, value, null);
                    } : (Action<T, object>)null,
                    ExportConverter = GetExportConverter(prop.PropertyType),
                    ImportConverter = GetImportConverter(prop.PropertyType)
                });
            }

            return result;
        }

        static Func<object, object> GetExportConverter(Type propertyType)
        {
            if (TypeInfo.IsNullableType(propertyType))
                return GetNullableConverter(GetExportConverter(Nullable.GetUnderlyingType(propertyType)));
            
            if (propertyType == typeof(Guid))
                return GuidExportConverter;           
            return null;
        }

        static Func<object, object> GetImportConverter(Type propertyType)
        {
            if (TypeInfo.IsNullableType(propertyType))
                return GetNullableConverter(GetImportConverter(Nullable.GetUnderlyingType(propertyType)));

            if (propertyType == typeof(DateTime))
                return DateTimeImportConverter;            
            if (propertyType == typeof(Guid))
                return GuidImportConverter;            

            return GetStandardConverter(propertyType);
        }

        private static Func<object, object> GetNullableConverter(Func<object, object> converter)
        {
            if (converter == null)
                return null;                    
            return (value) =>
            {
                if (value == null)
                    return null;
                return converter(value);
            };
        }

        static object DateTimeImportConverter(object o)
        {
            if (o is Double)
                return XlioUtil.ToDateTime((double)o);
            return Convert.ChangeType(o, typeof(DateTime), CultureInfo.InvariantCulture);
        }   

        static object GuidImportConverter(object o)
        {
            if (o is String)
                return new Guid((String)o);
            return o;
        }

        static object GuidExportConverter(object o)
        {
            if (o is Guid)
                return o.ToString();
            return o;
        }

        static object NullableGuidConverter(object o)
        {
            if (o == null)
                return null;
            return (Guid?)GuidImportConverter(o);
        }

        public static Func<object, object> GetStandardConverter(Type type)
        {
            return (value) => { return Convert.ChangeType(value, type, CultureInfo.InvariantCulture); };
        }
    }

    


}
