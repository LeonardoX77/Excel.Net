using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Reflection;
using System.Linq.Expressions;

namespace Utils
{
    public abstract class ReflectionTool<TSource>
    {
        private static List<string> FieldNames = null;

        /// <summary>
        /// Copy all values from Source Object to Destination Object
        /// </summary>
        /// <typeparam name="TDest">Destination type</typeparam>
        /// <param name="Source">Source Object</param>
        /// <param name="Destination">Destination Object</param>
        public static void CopyValues<TDest>(object Source, object Destination)
        {
            CopyValues<TDest>(null, Source, Destination);
        }

        /// <summary>
        /// Copy values from one instance class to other, specifying what field will be copied using Lambda Expressions
        /// </summary>
        /// <typeparam name="TDest">Destination Type</typeparam>
        /// <param name="Source">Source instance which contains data that will be copied</param>
        /// <param name="Destination">Target instance in which will be copied source data</param>
        /// <param name="SourceFieldList">Lambda expression wich specifyes what source fields will be copied
        /// Usage:
        /// ReflectionTool<Source_DataType>.CopyValues<Dest_DataType>(source, dest, c => new { c.Field1, c.Field2, ...});
        /// </param>
        /// <remarks>
        /// Source and Destination objects can be of different data types, but the Fields or porperties 
        /// that will be copied must be equal in both clases</remarks>
        public static void CopyValues<TDest>(object Source, object Destination, Expression<Func<TSource, object>> SourceFieldList)
        {
            List<string> Fields = GetFieldsFromLinQExpression(SourceFieldList);

            CopyValues<TDest>(Fields, Source, Destination);
        }

        public static List<string> GetFieldsFromLinQExpression(Expression<Func<TSource, object>> SourceFieldList)
        {
            List<string> Fields = new List<string>();
            foreach (MemberExpression item in (SourceFieldList.Body as NewExpression).Arguments)
            {
                Fields.Add(item.Member.Name);
            }

            return Fields;
        }

        private static void CopyValues<TDest>(List<string> FieldList, object Source, object Destination)
        {
            foreach (var propName in ReflectionTool<TSource>.getPropertyNames())
            {
                if ((FieldList == null || FieldList.Contains(propName))
                    && //check if destination object has the same propetyName
                    ReflectionTool<TDest>.ContainsPropertyName(propName))
                {
                    var value = ReflectionTool<TSource>.GetPropertyValue(Source, propName, GetDefaultValue(Source.GetType()));
                    ReflectionTool<TDest>.SetPropertyValue(Destination, propName, value);
                }
            }
        }

        public static object GetDefaultValue(Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }
            return null;
        }

        public static string GetPropertyName(Expression<Func<TSource, object>> SourceField)
        {
            var arguments = (SourceField.Body as NewExpression).Arguments;
            if (arguments.Count > 0)
                return ((MemberExpression)arguments[0]).Member.Name;
            else
                return "";
        }

        public static object GetPropertyValue(object src, string propName, object defaultValue)
        {
            object result = src.GetType().GetProperty(propName).GetValue(src, null) ?? defaultValue;
            return result;
        }
        public static void SetPropertyValue(object src, string propName, object Value)
        {
            src.GetType().GetProperty(propName).SetValue(src, Value);
        }

        public static Type GetPropertyType(string propName)
        {
            var propertyType = typeof(TSource).GetProperty(propName).PropertyType;

            if (propertyType.IsGenericType &&
                    propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                propertyType = propertyType.GetGenericArguments()[0];
            }

            return propertyType;
        }

        public static bool ContainsPropertyName(string propName)
        {
            return typeof(TSource).GetProperty(propName) != null;
        }

        public static List<string> getPropertyNames(string[] ExcludedFieldnames = null)
        {
            if (FieldNames == null)
            {
                FieldNames = new List<string>();

                //Public Fields 
                bool found = false;
                foreach (PropertyInfo prop in typeof(TSource).GetProperties())
                {
                    found = false;
                    if (ExcludedFieldnames != null)
                    {
                        foreach (var field in ExcludedFieldnames)
                        {
                            if (field == prop.Name)
                                found = true;
                        }
                    }

                    if (!found)
                        FieldNames.Add(prop.Name); //, prop.GetValue(this));
                    //Console.WriteLine("\t{0} {1} = {2}", prop.PropertyType.Name,
                    //    prop.Name, prop.GetValue(this, null));
                }


                // get single private Field -------------------------------------------------
                //FieldInfo FieldsExcluded = type.GetField("XLS_EXPORT_FIELD_NAMES_EXCLUDED", BindingFlags.NonPublic | BindingFlags.Static);
                //object value;
                //if (FieldsExcluded != null)
                //    value = FieldsExcluded.GetValue(null);

                //loop all private fields -------------------------------------------------
                //foreach (FieldInfo field in type.GetFields(BindingFlags.NonPublic | BindingFlags.Static))
                //{
                //    Console.WriteLine("\t{0} {1} = {2}", field.FieldType.Name,
                //        field.Name, field.GetValue(this));

                //    if (field.Name == "XLS_EXPORT_FIELD_NAMES_EXCLUDED")
                //    {
                //        //this.FieldNames.Add(prop.GetValue(this).ToString());
                //        FieldsExcluded = field.GetRawConstantValue().ToString(); //value of constants
                //    }
                //}

                //Public Fields -------------------------------------------------
                //foreach (PropertyDescriptor prop in TypeDescriptor.GetProperties(type))
                //{
                //        FieldNames.Add(prop.Name); //, prop.GetValue(this));
                //}

            }

            return FieldNames;
        }

        /// <summary>
        /// Returns a dictionary containing all class Properties and its values
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static Dictionary<string, object> GetFieldsAndValues(Object obj)
        {
            Dictionary<string, object> ret = new Dictionary<string, object>();
            foreach (var item in ReflectionTool<TSource>.getPropertyNames())
            {
                ret.Add(item, ReflectionTool<TSource>.GetPropertyValue(obj, item, ""));
            }

            return ret;
        }

    }
}