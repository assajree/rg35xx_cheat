using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace svvv.Helper
{
    public static class Extensions
    {

        public static void SetDefault(this TextBox txt, string defaultText)
        {
            if (String.IsNullOrWhiteSpace(txt.Text))
                txt.Text = defaultText;
        }

        public static int? ToIntOrNull(this string txt)
        {
            int result;
            if (Int32.TryParse(txt, out result))
                return result;
            else
                return null;
        }

        public static string NullIfEmpty(this string txt)
        {
            if(String.IsNullOrWhiteSpace(txt))
            {
                return null;
            }

            return txt;
        }

        public static string ToStringOrNull(this object obj)
        {
            if (obj==null)
            {
                return null;
            }

            return obj.ToString();
        }

        public static bool ToBoolean(this string obj,bool defaultValue)
        {
            if (Boolean.TryParse(obj, out var val))
                return val;
            else
                return defaultValue;
            
        }

        public static int ToInt32(this string obj, int defaultValue)
        {
            if (Int32.TryParse(obj, out var val))
                return val;
            else
                return defaultValue;

        }

        //public static Color ToColor(this string str)
        //{
        //    return System.Drawing.ColorTranslator.FromHtml(str);

        //}

      

        public static string Right(this string str, int length)
        {
            if (length >= str.Length)
                return str;

            return str.Substring(str.Length - length);
        }

        //public static String ToHexCode(this Color c)
        //{
        //    return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        //}

        //public static String ToRGB(this Color c)
        //{
        //    return "RGB(" + c.R.ToString() + "," + c.G.ToString() + "," + c.B.ToString() + ")";
        //}
    }
}
