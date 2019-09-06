using System;

namespace EasyNPOI.Attributes
{
    /// <summary>
    /// 普通模板占位符特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class PlaceholderAttribute : Attribute
    {
        public PlaceholderAttribute(string placeholderName)
        {
            PlaceholderName = placeholderName;
        }

        public string PlaceholderName { get; set; }

    }

    /// <summary>
    /// 表格模板占位符特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class PlaceholderGridAttribute : Attribute
    {
        public PlaceholderGridAttribute()
        {
        }

        public PlaceholderGridAttribute(string placeholderName)
        {
            PlaceholderName = placeholderName;
        }

        public string PlaceholderName { get; set; }

    }
}
