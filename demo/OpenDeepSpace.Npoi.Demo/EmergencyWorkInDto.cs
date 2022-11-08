using OpenDeepSpace.Npoi.Attributes;
using System.ComponentModel.DataAnnotations;

namespace OpenDeepSpace.Npoi.Demo
{
    public class EmergencyWorkerInDto
    {

        /// <summary>
        /// 响应等级Id
        /// </summary>
        public Guid ResponseLevelId { get; set; }

        /// <summary>
        /// 部门
        /// </summary>
        [ColumnNotNull(ErrorMessage = "第{row}行数据:部门不能为空")]
        [ExcelColumn(ColOrder = 1, ColName = "部门")]
        public string Department { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        [ColumnNotNull(ErrorMessage = "姓名不能为空")]
        [ExcelColumn(ColOrder = 2, ColName = "姓名")]
        public string Name { get; set; }


        /// <summary>
        /// 电话
        /// </summary>
        [ColumnNotNull(ErrorMessage = "手机号不能为空")]
        [ExcelColumn(ColOrder = 4, ColName = "手机号")]
        public string PhoneNumber { get; set; }


        /// <summary>
        /// 工号
        /// </summary>
        [Required(ErrorMessage = "工号不能为空")]
        [ColumnNotNull(ErrorMessage = "工号不能为空")]
        [ExcelColumn(ColOrder = 3, ColName = "工号")]
        public string JobNumber { get; set; }


        /// <summary>
        /// 备用电话
        /// </summary>
        [ExcelColumn(ColOrder = 5, ColName = "备用手机号")]
        public string AlternatePhoneNumber { get; set; }
    }
}
