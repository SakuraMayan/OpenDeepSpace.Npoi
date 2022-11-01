using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OpenDeepSpace.Npoi.Attributes;

namespace OpenDeepSpace.Npoi.Demo
{

    /// <summary>
    /// 泛指ExcelData数据输出对象
    /// </summary>
    [ExcelSheet(SheetName = "主数据")]
    public class ExcelDataOutDto
    {

        [ExcelColumn("唯一标识", 1, MergeColumn = true)]//通过ExcelColumn特性指定为需要输出到Excel的列,参数分别为列名,列排序
        /// <summary>
        /// Id
        /// </summary>
        public string Id { get; set; }
        //[NotNull(ErrorMessage = "姓名不能为空")]s
        //[MaxLength(MaxLength = 100, ErrorMessage = "姓名最大长度不能超过100个字符")]
        [ExcelColumn("姓名", 2, IsBaselineCol = true, MergeColumn = true)]
        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }

        [ExcelColumn("年龄", 3, MergeColumn = true)]
        /// <summary>
        /// 年龄
        /// </summary>
        public int Age { get; set; }

        [ExcelColumn("出生日期", 4, DatePattern = "yyyy-MM-dd HHmmss")]
        /// <summary>
        /// 出生日期
        /// </summary>
        public DateTime BirthDate { get; set; }

        /// <summary>
        /// 其他信息
        /// </summary>
        public string OtherInfo { get; set; }

        /// <summary>
        /// 关联数据
        /// </summary>
        [ExcelSheet]
        public List<ExcelRelationDataOutDto> ExcelRelationDataOutDtos { get; set; }
    }
}
