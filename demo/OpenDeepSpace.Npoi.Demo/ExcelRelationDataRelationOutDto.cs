using OpenDeepSpace.Npoi.Attributes;

namespace OpenDeepSpace.Npoi.Demo
{

    /// <summary>
    /// <see cref="ExcelRelationDataOutDto"/>的关联子数据
    /// </summary>
    [ExcelSheet(SheetName = "孙子数据")]
    public class ExcelRelationDataRelationOutDto
    {
        [ExcelColumn(ColOrder = 1, ColName = "名称", MergeColumn = true, IsBaselineCol = true)]
        public string Name { get; set; }
    }
}
