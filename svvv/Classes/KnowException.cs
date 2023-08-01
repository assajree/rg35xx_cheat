using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace svvv.Classes
{
    public class KnowException : Exception
    {
        public string Title { get; set; }
        public string Text { get; set; }


        public KnowException() : base()
        {
        }
        public KnowException(string message) : base(message)
        {
            Text = message;
        }

        public KnowException(string message, string title) : this(message)
        {
            Title = title;
        }

        public KnowException(string message, Exception innerException): base(message, innerException)
        {

        }
    }
}
