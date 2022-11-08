using System;
using System.Collections.Generic;
using System.Text;

namespace OpenDeepSpace.Npoi.Attributes
{
    /// <summary>
    /// 数据验证
    /// </summary>
    [AttributeUsage(AttributeTargets.Property,AllowMultiple =false)]
    public abstract class DataValidationAttribute:Attribute
    {
        /// <summary>
        /// 错误信息 错误信息中可以通过增加{row}占位提示错误在某一行
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 验证返回验证结果
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>
        public abstract DataValidationResult IsValid(object data);
    }
}
