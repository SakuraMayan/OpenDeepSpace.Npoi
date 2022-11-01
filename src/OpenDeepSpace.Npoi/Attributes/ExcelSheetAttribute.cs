using System;
using System.Collections.Generic;
using System.Text;

namespace OpenDeepSpace.Npoi.Attributes
{

    /// <summary>
    /// 指定Sheet的名称 或标注某属性对应的是Sheet
    /// </summary>
    [AttributeUsage(AttributeTargets.Class|AttributeTargets.Property)]
    public class ExcelSheetAttribute:Attribute
    {
        /// <summary>
        /// Sheet名称
        /// </summary>
        public string SheetName { get; set; }
    }
}
