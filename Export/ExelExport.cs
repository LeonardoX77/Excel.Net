using System;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;
using System.Net;

namespace Utils
{
    /// <summary>
    /// Utils for Export to Excel 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class ExelExport<T>
    {
        public static bool ExportExcel(List<T> dataList, string FileName, string FilePath, string[] ExcludedFieldnames = null)
        {
            StringBuilder sb = new StringBuilder();

            if (dataList.Count > 0)
            {
                string conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'";

                if (System.IO.File.Exists(FilePath))
                    System.IO.File.Delete(FilePath);

                using (OleDbConnection con = new OleDbConnection(conString))
                {
                    if (con.State == ConnectionState.Closed)
                        con.Open();

                    OleDbCommand cmd = new OleDbCommand(getXLSCreteScript(ExcludedFieldnames), con);
                    cmd.ExecuteNonQuery();

                    OleDbCommand cmdIns = new OleDbCommand(getXLSInsertScript(), con);

                    List<string> FieldList = ReflectionTool<T>.getPropertyNames(ExcludedFieldnames);
                    foreach (string fieldName in FieldList)
                    {
                        OleDbType paramType = GetOleDbType(ReflectionTool<T>.GetPropertyType(fieldName));
                        if (paramType == OleDbType.VarChar)
                            cmdIns.Parameters.Add("@" + fieldName, paramType, 4000);
                        else
                            cmdIns.Parameters.Add("@" + fieldName, paramType);
                    }

                    foreach (object field in dataList)
                    {
                        try
                        {
                            foreach (string fieldName in FieldList)
                            {
                                var type = ReflectionTool<T>.GetPropertyType(fieldName);
                                object defaultValue = "";
                                if (type.Equals(typeof(String)))
                                    defaultValue = (object)"";
                                else if(type.Equals(typeof(DateTime)))
                                    defaultValue = DBNull.Value; //new DateTime(); //(DateTime?)null;
                                else
                                    defaultValue = (object)0;

                                cmdIns.Parameters["@" + fieldName].Value = ReflectionTool<T>.GetPropertyValue(field, fieldName, defaultValue);
                            }
                            cmdIns.ExecuteNonQuery();
                        }
                        catch (Exception)
                        {
                            if (!System.Diagnostics.Debugger.IsAttached)
                            {
                                throw;
                            }
                        }
                    }

                    con.Close();
                }
            }

            return true;
        }

        public static string getXLSCreteScript(string[] ExcludedFieldnames = null)
        {
            string result = "";
            foreach (var item in ReflectionTool<T>.getPropertyNames(ExcludedFieldnames))
            {
                result += (result.Length == 0 ? "" : ", \r\n") +
                    "[" + item + "] " + GetOleDbTypeString(ReflectionTool<T>.GetPropertyType(item)).ToString().ToLower();
            }

            result = "CREATE TABLE " + typeof(T).Name + "s (" + result + ")";
            return result;
        }

        public static string getXLSInsertScript(string[] ExcludedFieldnames = null)
        {
            string result = "";
            string fields = "", values = "";

            foreach (var item in ReflectionTool<T>.getPropertyNames(ExcludedFieldnames))
            {
                fields += (fields.Length == 0 ? "" : ", ") + "[" + item + "] ";
                values += (values.Length == 0 ? "" : ", ") + "@" + item;
            }

            result = "INSERT INTO " + typeof(T).Name + "s (" + fields + ") VALUES (" + values + ")";
            return result;

        }

        private static string GetOleDbTypeString(Type sysType)
        {
            string result = "";
            if (sysType.Equals(typeof(Double)))
                result = "float";

            else if (sysType.Equals(typeof(String)))
                result = "text";

            else if (sysType.Equals(typeof(DateTime)))
                result = "datetime";

            else if (sysType.Equals(typeof(Single)))
                result = "real";

            else if (sysType.Equals(typeof(int)) || sysType.Equals(typeof(Int32)) || (sysType.BaseType != null && sysType.BaseType.Equals(typeof(System.Enum))))
                result = "int";

            else if (sysType.Equals(typeof(bool)) || sysType.Equals(typeof(Boolean)))
                result = "bit";

            else
                throw new TypeLoadException(string.Format("Can not load CLR Type from {0}", sysType.Name));

            return result;
            //Mappings.Add("bigint", typeof(Int64));
            //Mappings.Add("binary", typeof(Byte[]));
            //Mappings.Add("bit", typeof(Boolean));
            //Mappings.Add("char", typeof(String));
            //Mappings.Add("date", typeof(DateTime));
            //Mappings.Add("datetime", typeof(DateTime));
            //Mappings.Add("datetime2", typeof(DateTime));
            //Mappings.Add("datetimeoffset", typeof(DateTimeOffset));
            //Mappings.Add("decimal", typeof(Decimal));
            //Mappings.Add("float", typeof(Double));
            //Mappings.Add("image", typeof(Byte[]));
            //Mappings.Add("int", typeof(Int32));
            //Mappings.Add("money", typeof(Decimal));
            //Mappings.Add("nchar", typeof(String));
            //Mappings.Add("ntext", typeof(String));
            //Mappings.Add("numeric", typeof(Decimal));
            //Mappings.Add("nvarchar", typeof(String));
            //Mappings.Add("real", typeof(Single));
            //Mappings.Add("rowversion", typeof(Byte[]));
            //Mappings.Add("smalldatetime", typeof(DateTime));
            //Mappings.Add("smallint", typeof(Int16));
            //Mappings.Add("smallmoney", typeof(Decimal));
            //Mappings.Add("text", typeof(String));
            //Mappings.Add("time", typeof(TimeSpan));
            //Mappings.Add("timestamp", typeof(Byte[]));
            //Mappings.Add("tinyint", typeof(Byte));
            //Mappings.Add("uniqueidentifier", typeof(Guid));
            //Mappings.Add("varbinary", typeof(Byte[]));
            //Mappings.Add("varchar", typeof(String));

        }

        private static OleDbType GetOleDbType(Type sysType)
        {
            if (sysType.Equals(typeof(String)))
                return OleDbType.LongVarChar;

            else if (sysType.Equals(typeof(int)) || sysType.BaseType.Equals(typeof(System.Enum)))
                return OleDbType.Integer;

            else if (sysType.Equals(typeof(Boolean)))
                return OleDbType.Boolean;

            else if (sysType.Equals(typeof(DateTime)))
                return OleDbType.Date;

            else if (sysType.Equals(typeof(Char)))
                return OleDbType.Char;

            else if (sysType.Equals(typeof(Decimal)))
                return OleDbType.Decimal;

            else if (sysType.Equals(typeof(Double)))
                return OleDbType.Double;

            else if (sysType.Equals(typeof(Single)))
                return OleDbType.Single;

            else if (sysType.Equals(typeof(Byte)))
                return OleDbType.Binary;

            else if (sysType.Equals(typeof(Guid)))
                return OleDbType.Guid;

            else
                throw new TypeLoadException(string.Format("Can not load CLR Type from {0}", sysType.Name));
        }





    }

}