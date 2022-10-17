using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDeepSpace.Npoi.Attributes
{

    /// <summary>
    /// Excel列信息 名称和序号
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelColumnAttribute:Attribute
    {
        private string colName;//Excel中的类名

        private int colOrder;//规定列的顺序

        /// <summary>
        /// 合并相同数据行
        /// </summary>
        private bool mergeRow;//合并相同数据行

        /// <summary>
        /// 合并相同数据列
        /// </summary>
        private bool mergeColumn;//合同相同数据列

        private bool isBaselineCol;//是否基准列 基准列为一个 一般基准列为唯一标识列
        private bool isCalculateCol;//是否计算列

        private string datePattern;//时间格式化

        public ExcelColumnAttribute() { }

        public ExcelColumnAttribute(string colName) { this.colName = colName; }

        public ExcelColumnAttribute(string colName, int colOrder) { this.colName = colName; this.colOrder = colOrder; }

        public string ColName { get => colName; set => colName = value; }
        public int ColOrder { get => colOrder; set => colOrder = value; }
        public bool MergeRow { get => mergeRow; set => mergeRow = value; }
        public bool MergeColumn { get => mergeColumn; set => mergeColumn = value; }
        public bool IsBaselineCol { get => isBaselineCol; set => isBaselineCol = value; }
        public bool IsCalculateCol { get => isCalculateCol; set => isCalculateCol = value; }
        public string DatePattern { get => datePattern; set => datePattern = value; }
    }
}
