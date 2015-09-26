using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace ExtensionMethods
{

    internal static partial class ExtensionMethodsCLR
    {

        # region Non Portable

#if !PORTABLE
        public static SecureString ToSecureString(this string str)
        {
            SecureString password = new SecureString();
            foreach (char c in str)
            {
                password.AppendChar(c);
            }
            // lock the password down
            password.MakeReadOnly();
            return password;

        }

        public static Guid ParseAsGuid(this object obj, out string error)
        {
            error = null;
            Guid guid;
            if (Guid.TryParse(obj.ToString(), out guid)) return guid;

            try
            {
                Guid.Parse(obj.ToString());
            }
            catch (Exception e)
            {
                error = e.Message;
            }
            return Guid.Empty;
        }

        public static void WriteToFile<T>(this T obj, string filePath)
        {
            using (FileStream file = new FileStream(filePath, FileMode.Create))
            {
                (new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()).Serialize(file, obj);
            }
        }

        public static T ReadFromFile<T>(this T obj, string filePath)
        {
            using (FileStream file = new FileStream(filePath, FileMode.Open))
            {
                return (T)(new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()).Deserialize(file);
            }
        }

        public static bool SerializeTo<T>(this object obj, string filepath)
        {
            bool res = false;
            try
            {
                TextWriter writer = new StreamWriter(filepath, append: false);
                XmlSerializer ser = new XmlSerializer(typeof(T));
                ser.Serialize(writer, obj);
                writer.Flush();
                writer.Close();
                res = true;
            }
            catch (Exception)
            { }

            return res;
        }
        public static T DeserializeFrom<T>(this string s)
        {
            T res = default(T);
            try
            {
                TextReader reader = new StreamReader(s);
                XmlSerializer ser = new XmlSerializer(typeof(T));
                res = (T)ser.Deserialize(reader);
            }
            catch (Exception)
            { }
            //if (File.Exists(s))
            //{            }

            return res;
        }

        public static string EscapeInvalidFileNameChars(this string s, string separator = "")
        {
            string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
            return r.Replace(s, separator);
        }

        public static string ReadTextFromFile(this string filepath)
        {
            try
            {
                using (TextReader tr = new StreamReader(filepath))
                {
                    return tr.ReadToEnd();
                }

            }
            catch (Exception)
            {


            }

            return null;

        }

        public static bool WriteTextToFile(this string content, string filepath, bool append = false)
        {
            try
            {
                using (TextWriter tr = new StreamWriter(filepath, append))
                {
                    tr.Write(content);
                    return true;
                }

            }
            catch (Exception)
            {


            }

            return false;

        }

#endif

        #endregion




        //public static string RemoveLeadingNewLine(this string str)
        //{

        //    if (str.StartsWith("\r\n"))
        //    {
        //        return RemoveLeadingNewLine(str.Remove(0, 2));
        //    }
        //    return str;
        //}


        public static bool? ParseBool(this object obj, bool indexOfTrue = true, bool nullable = true)
        {
            if (obj == null) return null;

            return obj.ToString().ParseBool(indexOfTrue, nullable);
        }
        public static bool? ParseBool(this string str, bool indexOfTrue = true, bool nullable = true)
        {
            bool outnull;
            if (nullable && string.IsNullOrEmpty(str)) return null;
            if (bool.TryParse(str, out outnull))
                return bool.Parse(str);
            else if (str.Equals("$true", StringComparison.OrdinalIgnoreCase))
                return true;
            else if (indexOfTrue && str.IndexOf("true", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            else if (str.Equals("0"))
                return false;
            else if (str.Equals("1"))
                return true;
            else if (nullable)
                return null;
            else
                return false;
        }

        public static bool IsNullOrEmpty(this string str)
        {
            return string.IsNullOrEmpty(str);
        }

        public static int ContainsCount(this string str, string containsStr, bool ignoreCase = true)
        {
            int i = 0;
            if (ignoreCase)
            {
                str = str.ToLower();
                containsStr = containsStr.ToString();
            }

            while (str.Contains(containsStr))
            {
                str = str.Substring(str.IndexOf(containsStr) + containsStr.Length);
                i++;
            }

            return i;
        }

        public static int ParseAsInt(this object obj, out string error)
        {
            error = null;
            int i;
            if (int.TryParse(obj.ToString(), out i)) return i;

            try
            {
                int.Parse(obj.ToString());
            }
            catch (Exception e)
            {
                error = e.Message;
            }
            return 0;
        }

        /// <summary>
        /// Parse the array of object as string separated by delemeter, default delemeter is ",".
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="error"></param>
        /// <param name="delemeter"></param>
        /// <returns></returns>
        public static string ParseAsArrayString(this object obj, out string error, string delemeter = ",")
        {
            error = null;
            string s = string.Empty;

            object[] arObj = obj as object[];
            if (arObj != null)
            {
                foreach (var item in arObj)
                {
                    s += item.ToString() + delemeter;
                }
                int lastindex = s.LastIndexOf(delemeter);
                s = lastindex >= 0 ? s.Remove(lastindex) : s;
                return s;
            }
            error = "not an array";
            return obj.ToString();
        }

        public static double ParseAsDouble(this object obj, out string error)
        {
            error = null;
            double d;
            if (double.TryParse(obj.ToString(), out d)) return d;

            try
            {
                double.Parse(obj.ToString());
            }
            catch (Exception e)
            {
                error = e.Message;
            }
            return double.NaN;
        }

        public static float ParseAsFloat(this object obj, out string error)
        {
            error = null;
            float f;
            if (float.TryParse(obj.ToString(), out f)) return f;

            try
            {
                float.Parse(obj.ToString());
            }
            catch (Exception e)
            {
                error = e.Message;
            }
            return float.NaN;
        }

        public static DateTime ParseAsDateTime(this object obj, out string error)
        {
            error = null;
            if (obj is DateTime) return DateTime.Parse(obj.ToString(), (IFormatProvider)CultureInfo.InvariantCulture);

            if (obj is DateTimeOffset) return DateTimeOffset.Parse(obj.ToString(), (IFormatProvider)CultureInfo.InvariantCulture).DateTime;

            //error = null;
            //DateTime dt;
            //if (DateTime.TryParse(obj.ToString(), out dt)) return dt;            
            //try
            //{
            //    DateTime.Parse(obj.ToString());
            //}
            //catch (Exception e)
            //{
            //    error = e.Message;
            //}

            error = "not a valid DateTime or DateTimeOffset type. Default value retuened";
            return default(DateTime);
        }



        //public static bool TrySerializeToString(this object obj, Type type, out string s)
        //{
        //    bool result = false;
        //    s = string.Empty;
        //    try
        //    {
        //        using (StringWriter writer = new StringWriter())
        //        {
        //            XmlSerializer ser = new XmlSerializer(type);
        //            ser.Serialize(writer, obj);
        //            s = writer.ToString();
        //            result = true;
        //        }
        //    }
        //    catch (Exception)
        //    { s = null; }

        //    return result;

        //}
        //public static bool TryDeserializeToObject(this string data, Type type, out object obj)
        //{
        //    bool result = false;
        //    obj = null;
        //    try
        //    {
        //        using (StringReader reader = new StringReader(data))
        //        {
        //            XmlSerializer ser = new XmlSerializer(type);
        //            obj = ser.Deserialize(reader);
        //            result = true;
        //        }
        //    }
        //    catch (Exception)
        //    { obj = null; }

        //    return result;
        //}

        public static bool TrySerializeToString<T>(this object obj, out string s)
        {
            bool result = false;
            s = string.Empty;
            try
            {
                using (StringWriter writer = new StringWriter())
                {
                    XmlSerializer ser = new XmlSerializer(typeof(T));
                    ser.Serialize(writer, obj);
                    s = writer.ToString();
                    result = true;
                }
            }
            catch (Exception)
            { s = null; }

            return result;

        }
        public static bool TryDeserializeToObject<T>(this string data, out T obj)
        {
            bool result = false;
            obj = default(T);
            try
            {
                using (StringReader reader = new StringReader(data))
                {
                    XmlSerializer ser = new XmlSerializer(typeof(T));
                    obj = (T)ser.Deserialize(reader);
                    result = true;
                }
            }
            catch (Exception)
            { obj = default(T); }

            return result;
        }

        public static List<T> ToSingleItemList<T>(this T t)
        {
            List<T> _r = new List<T>();
            _r.Add(t);
            return _r;
        }

        public static string ToHTMLFriendlyString(this string s)
        {
            return s.Replace("\r\n", "<br />").Replace("\n", "<br />");
        }



    }
}
