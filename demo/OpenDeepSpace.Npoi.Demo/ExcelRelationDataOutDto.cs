using OpenDeepSpace.Npoi.Attributes;

namespace OpenDeepSpace.Npoi.Demo
{

    [ExcelSheet(SheetName = "关联数据")]
    public class ExcelRelationDataOutDto
    {
        [ExcelColumn(ColOrder =1,ColName ="名称",MergeColumn =true,IsBaselineCol =true)]
        public string Name { get; set; }
    }
}
