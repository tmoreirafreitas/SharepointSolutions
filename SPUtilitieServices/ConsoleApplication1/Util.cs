using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace SPUtilitieServicesSolution
{
    public static class Util
    {
        public static T Cast<T>(object obj, T type)
        {
            return (T)obj;
        }

        public static void AddProperty(ExpandoObject expando, string propertyName, object propertyValue)
        {
            // ExpandoObject supports IDictionary so we can extend it like this
            var expandoDict = expando as IDictionary<string, object>;
            if (expandoDict.ContainsKey(propertyName))
                expandoDict[propertyName] = propertyValue;
            else
                expandoDict.Add(propertyName, propertyValue);
        }
        public static bool IsPropertyExist(dynamic settings, string name)
        {
            if (settings is ExpandoObject)
                return ((IDictionary<string, object>)settings).ContainsKey(name);

            return settings.GetType().GetProperty(name) != null;
        }

        /* -- Format for Date Only Field -- */
        public static String ToSPDate(String strDt)
        {
            if (strDt == String.Empty)
                return strDt;
            else
                return (Convert.ToDateTime(strDt)).ToString("yyyy-MM-dd");
        }

        /* -- Format for DateTime Field -- */
        public static String ToSPDateTime(String strDt)
        {
            if (strDt == String.Empty)
                return strDt;
            else
                return (Convert.ToDateTime(strDt)).ToString("yyyy-MM-ddTHH:mm:ssZ");
        }

        public static DataTable GetDataTableFromListDynamic(IList<IDictionary<string, object>> items)
        {
            DataTable table = new DataTable();
            try
            {
                if (items != null && items.Count > 0)
                {
                    var obj = items[0];
                    foreach (var key in obj.Keys)
                        table.Columns.Add(key);

                    foreach (var item in items)
                    {
                        DataRow dr = table.NewRow();
                        foreach (var columName in item.Keys)
                            dr.SetField(columName, item[columName]);
                        table.Rows.Add(dr);
                    }
                }
                return table;
            }
            catch (ArgumentException ex)
            {
                throw ex;
            }
        }

        public static DataTable GetDataTableFromListDictionary(IList<IDictionary<string, string>> items)
        {
            DataTable table = new DataTable();
            try
            {
                if (items != null && items.Count > 0)
                {
                    var obj = items[0];
                    foreach (var key in obj.Keys)
                        table.Columns.Add(key);

                    foreach (var item in items)
                    {
                        DataRow dr = table.NewRow();
                        foreach (var columName in item.Keys)
                            dr.SetField(columName, item[columName]);
                        table.Rows.Add(dr);
                    }
                }
                return table;
            }
            catch (ArgumentException ex)
            {
                throw ex;
            }
        }

        public static void CreateCSV(DataTable dataTable, string filePath, string delimiter = ";")
        {
            if (!Directory.Exists(Path.GetDirectoryName(filePath)))
                throw new DirectoryNotFoundException($"Destination folder not found: {filePath}");

            var columns = dataTable.Columns.Cast<DataColumn>().ToArray();

            var lines = (new[] { string.Join(delimiter, columns.Select(c => c.ColumnName)) })
              .Union(dataTable.Rows.Cast<DataRow>().Select(row => string.Join(delimiter, columns.Select(c => row[c]))));
            File.WriteAllLines(filePath, lines, Encoding.UTF8);            
        }
    }
}
