using System;
using System.Collections.Generic;
using System.Text;

namespace OpenDeepSpace.Npoi
{
    /// <summary>
    /// 数据验证结果
    /// </summary>
    public class DataValidationResult
    {

        public static readonly DataValidationResult Success=null;

        public DataValidationResult(string errorMessage)
        {
            ErrorMessage = errorMessage;
        }

        public string ErrorMessage { get; set; }

    
    }
}
