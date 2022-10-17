using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OpenDeepSpace.Npoi
{

    /// <summary>
    /// Excel列对象
    /// </summary>
    public class ExcelColumn : IComparable<ExcelColumn>
    {

        private string colName;//列名
        private int colOrder;//序号
        private bool mergeRow;//是否合并相同数据行
        private bool mergeCol;//是否合并相同数据列
        private string propertyName;//列名特性下所对应的属性名
        private Type propertyType;//列名特性下对应的属性类型
        private PropertyInfo propertyInfo;//属性信息
        private bool isBaselineCol;//是否基准列 基准列为一个 一般基准列为唯一标识列
        private bool isCalculateCol;//是否计算列
        private string datePattern;//时间格式化

        public ExcelColumn()
        {
           
        }


        public ExcelColumn(string colName, int colOrder, string propertyName, Type propertyType)
        {

            this.ColName = colName;
            this.ColOrder = colOrder;
            this.PropertyName = propertyName;
            this.PropertyType = propertyType;
        }

        public string ColName { get => colName; set => colName = value; }
        public int ColOrder { get => colOrder; set => colOrder = value; }
        public string PropertyName { get => propertyName; set => propertyName = value; }
        public Type PropertyType { get => propertyType; set => propertyType = value; }
        public bool MergeRow { get => mergeRow; set => mergeRow = value; }
        public bool MergeCol { get => mergeCol; set => mergeCol = value; }
        public bool IsBaselineCol { get => isBaselineCol; set => isBaselineCol = value; }
        public bool IsCalculateCol { get => isCalculateCol; set => isCalculateCol = value; }
        public string DatePattern { get => datePattern; set => datePattern = value; }

        public PropertyInfo PropertyInfo { get => propertyInfo; set => propertyInfo = value; }

        public int CompareTo(ExcelColumn other)
        {

            return ColOrder > other.ColOrder ? 1 : (ColOrder < other.ColOrder) ? -1 : 0;
        }
    }
}
