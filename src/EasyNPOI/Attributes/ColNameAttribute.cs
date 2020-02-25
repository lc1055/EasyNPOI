using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyNPOI.Attributes
{
    /// <summary>
    /// Excel列名
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColNameAttribute : Attribute
    {
        public ColNameAttribute(string colName)
        {
            ColName = colName;
        }

        public string ColName { get; set; }
    }
}
