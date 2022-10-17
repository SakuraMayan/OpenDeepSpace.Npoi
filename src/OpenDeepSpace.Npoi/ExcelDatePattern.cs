using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDeepSpace.Npoi
{
	/// <summary>
	/// 全局时间格式化设置
	/// </summary>
    public class ExcelDatePattern
    {
        public ExcelDatePattern()
        {

        }

		private string datePattern = "yyyy-MM-dd";

		public string getDatePattern()
		{
			return datePattern;
		}

		public void setDatePattern(string datePattern)
		{
			this.datePattern = datePattern;
		}
	}
}
