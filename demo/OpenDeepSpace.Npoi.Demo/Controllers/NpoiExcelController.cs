using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace OpenDeepSpace.Npoi.Demo.Controllers
{
    /// <summary>
    /// Npoi对Excel操作控制器
    /// </summary>
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class NpoiExcelController : ControllerBase
    {

        public NpoiExcelController()
        {

        }

        List<ExcelDataOutDto> excelDataOutDtos = new List<ExcelDataOutDto>() {

                new ExcelDataOutDto(){
                    Id=Guid.NewGuid().ToString(),
                    Name="小张",
                    Age=20,
                    BirthDate=DateTime.Today.AddYears(-20),
                    OtherInfo="我是小张，我不嚣张"
                },
                new ExcelDataOutDto(){
                    Id="b66f90b7-178b-11ed-98e5-00155d562808",
                    Name="小李",
                    Age=24,
                    BirthDate=DateTime.Today.AddYears(-24),
                    OtherInfo="面对疾风吧"
                },
                new ExcelDataOutDto(){
                    Id="b66f90b7-178b-11ed-98e5-00155d562808",
                    Name="小李",
                    Age=24,
                    BirthDate=DateTime.Today.AddYears(-24),
                    OtherInfo="面对疾风吧八八八"
                },
                /*new ExcelDataOutDto(){
                    Id=Guid.NewGuid().ToString(),
                    Name="小红",
                    Age=24,
                    BirthDate=DateTime.Today.AddYears(-24),
                    OtherInfo="面对疾风吧八八八"
                },*/
                new ExcelDataOutDto(){
                    Id="b66f90b7-178b-11ed-98e5-00155d5s2808",
                    Name="小吴",
                    Age=27,
                    BirthDate=DateTime.Today.AddYears(-27),
                    OtherInfo="哈哈哈笑而不语"
                },
                new ExcelDataOutDto(){
                    Id="b66f90b7-178b-11ed-98e5-00155d5s2808",
                    Name="小吴",
                    Age=27,
                    BirthDate=DateTime.Today.AddYears(-27),
                    OtherInfo="哈哈哈笑笑笑"
                },
                new ExcelDataOutDto(){
                    Id="b66f90b7-178b-11ed-98e5-00155d5s2809",
                    Name="",
                    Age=27,
                    BirthDate=DateTime.Today.AddYears(-27),
                    OtherInfo="哈哈哈笑笑笑"
                },
            };

        /// <summary>
        /// 无模板导出数据到Excel到指定路径
        /// </summary>
        [HttpGet]
        [HttpPost]//都未指定template将合并
        //[HttpPost("postexcel")]
        public void ExportObjectToExcelToPath()
        {
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.exportObjectToExcel($@"D:\excelexport\{Guid.NewGuid().ToString("N")}.xlsx", excelDataOutDtos);
        }

        /// <summary>
        /// 无模板导出数据到Excel到指定路径
        /// 不带序号
        /// </summary>
        [HttpGet]
        public void ExportObjectToExcelToPathNoOrder()
        {
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.setIsSetOrder(false);
            ExcelHandle.getExcelDatePattern().setDatePattern("yyyy/MM/dd");//指定日期的格式化模式
            ExcelHandle.exportObjectToExcel($@"D:\excelexport\{Guid.NewGuid().ToString("N")}.xlsx", excelDataOutDtos);
        }

        /// <summary>
        /// 无模板导出数据到Excel到指定路径
        /// 指定导出的列
        /// </summary>
        [HttpGet]
        public void ExportObjectToExcelToPathAndColMaps()
        {
            //指定导出列 默认只要写了ExcelColumn特性的都会导出，这里可以对ExcelColumn进行进一步筛选
            List<string> colNames = new List<string>();
            colNames.Add("姓名");
            colNames.Add("出生日期");
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.exportObjectToExcel($@"D:\excelexport\{Guid.NewGuid().ToString("N")}.xlsx", excelDataOutDtos, colNames);
        }


        /// <summary>
        /// 无模板导出数据到excel并返回流文件供下载
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<FileResult> ExportObjectToExcelToStream()
        {
            ExcelHandle ExcelHandle = new ExcelHandle();
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();

            List<ExcelDataOutDto> excelDataOutDtosOver = new List<ExcelDataOutDto>();
            for (int i = 0; i < 10000; i++)
            {
                excelDataOutDtosOver.AddRange(excelDataOutDtos);
            }

            ExcelHandle.exportObjectToExcel(npoiMemoryStream, excelDataOutDtosOver);
            npoiMemoryStream.Seek(0, SeekOrigin.Begin);
            return File(npoiMemoryStream, "application/stream", $"{Guid.NewGuid()}.xlsx");
        }

        /// <summary>
        /// 无模板导出数据到excel并返回流文件供下载
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<FileResult> ExportObjectToExcelToStreamAndColMaps()
        {
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();

            List<string> colNames = new List<string>();
            colNames.Add("姓名");
            colNames.Add("出生日期");
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.exportObjectToExcel(npoiMemoryStream, excelDataOutDtos, colNames);
            npoiMemoryStream.Seek(0, SeekOrigin.Begin);

            return File(npoiMemoryStream, "application/stream", $"{Guid.NewGuid()}.xlsx");
        }


        /// <summary>
        /// 指定模板导出数据到Excel 无头部底部
        /// 模板在工程下面
        /// </summary>
        [HttpGet]
        public async Task<FileResult> ExportObjectToExcelByTemplateInClassPathToPath()
        {
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);//设置excel的列数 不论是否设置序号列需要算上即+1
            ExcelHandle.exportObjectToExcelByTemplate(Path.Combine(@"ExcelTemplates/导出的数据模板无头部底部.xlsx"), $@"D:\excelexporttemplate\指定路径模板{Guid.NewGuid().ToString("N")}.xlsx", excelDataOutDtos);

            ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);//设置excel的列数 不论是否设置序号列需要算上即+1
            //导出到流文件
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();

            ExcelHandle.exportObjectToExcelByTemplate(Path.Combine(@"ExcelTemplates/导出的数据模板无头部底部.xlsx"), npoiMemoryStream, excelDataOutDtos);

            npoiMemoryStream.Seek(0, SeekOrigin.Begin);

            return File(npoiMemoryStream, "application/stream", $"模板流文件无头部底部{Guid.NewGuid()}.xlsx");
        }

        /// <summary>
        /// 指定模板导出数据到Excel 无头部底部内容
        /// 指定模板所在路径
        /// </summary>
        [HttpGet]
        public async Task<FileResult> ExportObjectToExcelByTemplateInPathToPath()
        {
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.setIsClasspath(false);//模板不在工程目录下
            ExcelHandle.setIsSetOrder(false);//不设置序号
            ExcelHandle.getExcelTemplate().setExcelColNums(4);//设置excel的列数 不论是否设置序号列需要算上即+1 如果前面有空的列需要加上空的列数+2
            ExcelHandle.getExcelTemplate().setDataStartColNum(2);
            ExcelHandle.exportObjectToExcelByTemplate(@"C:\Users\lenovo\Desktop\导出的数据模板无头部底部.xlsx", $@"D:\excelexporttemplate\指定路径模板{Guid.NewGuid().ToString("N")}.xlsx", excelDataOutDtos);


            ExcelHandle = new ExcelHandle();
            ExcelHandle.setIsClasspath(false);//模板不在工程目录下
            ExcelHandle.setIsSetOrder(false);//不设置序号
            ExcelHandle.getExcelTemplate().setExcelColNums(4);//设置excel的列数 不论是否设置序号列需要算上即+1 如果前面有空的列需要加上空的列数+2
            ExcelHandle.getExcelTemplate().setDataStartColNum(2);
            //导出到流文件
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();

            ExcelHandle.exportObjectToExcelByTemplate(@"C:\Users\lenovo\Desktop\导出的数据模板无头部底部.xlsx", npoiMemoryStream, excelDataOutDtos);

            npoiMemoryStream.Seek(0, SeekOrigin.Begin);

            return File(npoiMemoryStream, "application/stream", $"模板流文件无头部底部{Guid.NewGuid()}.xlsx");
        }

        /// <summary>
        /// 根据模板文件流导出数据到excel 
        /// </summary>
        /// <param name="templateFormFile">模板文件</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<FileResult> ExportObjectToExcelByTemplate(IFormFile templateFormFile)
        {
            var streamone = templateFormFile.OpenReadStream();

            var streamTwo = new MemoryStream();
            streamone.CopyTo(streamTwo);
            streamone.Seek(0, SeekOrigin.Begin);
            streamTwo.Seek(0, SeekOrigin.Begin);

            var streamThree = new MemoryStream();
            streamone.CopyTo(streamThree);
            streamThree.Seek(0, SeekOrigin.Begin);

            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);
            ExcelHandle.getExcelTemplate().setDataStartColNum(2);




            ExcelHandle.exportObjectToExcelByTemplate(streamTwo, $@"D:\excelexportstreamtemplate\指定路径模板{Guid.NewGuid().ToString("N")}.xlsx", excelDataOutDtos);



            //导出到流
            ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);
            ExcelHandle.getExcelTemplate().setDataStartColNum(2);
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();
            ExcelHandle.exportObjectToExcelByTemplate(streamThree, npoiMemoryStream, excelDataOutDtos);
            npoiMemoryStream.Seek(0, SeekOrigin.Begin);

            return File(npoiMemoryStream, "application/stream", $"模板流文件无头部底部{Guid.NewGuid()}.xlsx");
        }

        /// <summary>
        /// 导出文件到excel包含头部底部
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<FileResult> ExportObjectToExcelIncludeHeaderAndFooter()
        {
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);
            Dictionary<string, object> headerAndFooterMaps = new Dictionary<string, object>();
            headerAndFooterMaps.Add("datetime", DateTime.Now);
            headerAndFooterMaps.Add("exporter", "OpenDeepSpace管理员");
            headerAndFooterMaps.Add("count", excelDataOutDtos.Count);
            ExcelHandle.exportObjectToExcelByTemplate(@"C:\Users\lenovo\Desktop\导出的数据模板含头部底部.xlsx", $@"D:\excelexporthftemplate\指定路径带头部底部模板{Guid.NewGuid().ToString("N")}.xlsx", headerAndFooterMaps, excelDataOutDtos);

            ExcelHandle = new ExcelHandle();
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);
            ExcelHandle.exportObjectToExcelByTemplate(@"C:\Users\lenovo\Desktop\导出的数据模板含头部底部.xlsx", npoiMemoryStream, headerAndFooterMaps, excelDataOutDtos);
            npoiMemoryStream.Seek(0, SeekOrigin.Begin);

            return File(npoiMemoryStream, "application/stream", $"模板流文件有头部底部{Guid.NewGuid()}.xlsx");
        }

        /// <summary>
        /// 根据模板文件流导出数据到excel 
        /// </summary>
        /// <param name="templateFormFile">模板文件</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<FileResult> ExportObjectToExcelByTemplateIncludeHeaderAndFooter(IFormFile templateFormFile)
        {
            var streamone = templateFormFile.OpenReadStream();

            var streamTwo = new MemoryStream();
            streamone.CopyTo(streamTwo);
            streamone.Seek(0, SeekOrigin.Begin);
            streamTwo.Seek(0, SeekOrigin.Begin);

            var streamThree = new MemoryStream();
            streamone.CopyTo(streamThree);
            streamThree.Seek(0, SeekOrigin.Begin);

            Dictionary<string, object> headerAndFooterMaps = new Dictionary<string, object>();
            headerAndFooterMaps.Add("datetime", DateTime.Now);
            headerAndFooterMaps.Add("exporter", "OpenDeepSpace管理员");
            headerAndFooterMaps.Add("count", excelDataOutDtos.Count);

            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);




            ExcelHandle.exportObjectToExcelByTemplate(streamTwo, $@"D:\excelexporthftemplate\指定路径模板{Guid.NewGuid().ToString("N")}.xlsx", headerAndFooterMaps, excelDataOutDtos);



            //导出到流
            ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();
            ExcelHandle.exportObjectToExcelByTemplate(streamThree, npoiMemoryStream, headerAndFooterMaps, excelDataOutDtos);
            npoiMemoryStream.Seek(0, SeekOrigin.Begin);

            return File(npoiMemoryStream, "application/stream", $"模板流文件有头部底部{Guid.NewGuid()}.xlsx");
        }


        /// <summary>
        /// 指定模板导出数据到Excel 无头部底部内容
        /// 进一步选择列
        /// 指定模板所在路径
        /// </summary>
        [HttpGet]
        public async Task<FileResult> ExportObjectToExcelByTemplateInPathToPathSelect()
        {
            //选择需要导出的列
            List<string> colNames = new List<string>();
            colNames.Add("姓名");
            colNames.Add("出生日期");
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.setIsClasspath(false);//模板不在工程目录下
            ExcelHandle.setIsSetOrder(false);//不设置序号
            ExcelHandle.getExcelTemplate().setExcelColNums(2);//设置excel的列数 不论是否设置序号列需要算上即+1 如果前面有空的列需要加上空的列数+2
            ExcelHandle.getExcelTemplate().setDataStartColNum(2);
            ExcelHandle.exportObjectToExcelByTemplate(@"C:\Users\lenovo\Desktop\导出的数据模板无头部底部.xlsx", $@"D:\excelexporttemplate\指定路径模板{Guid.NewGuid().ToString("N")}.xlsx", null, excelDataOutDtos, colNames);


            ExcelHandle = new ExcelHandle();
            ExcelHandle.setIsClasspath(false);//模板不在工程目录下
            ExcelHandle.setIsSetOrder(false);//不设置序号
            ExcelHandle.getExcelTemplate().setExcelColNums(2);//设置excel的列数 不论是否设置序号列需要算上即+1 如果前面有空的列需要加上空的列数+2
            ExcelHandle.getExcelTemplate().setDataStartColNum(2);
            //导出到流文件
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();

            ExcelHandle.exportObjectToExcelByTemplate(@"C:\Users\lenovo\Desktop\导出的数据模板无头部底部.xlsx", npoiMemoryStream, null, excelDataOutDtos, colNames);

            npoiMemoryStream.Seek(0, SeekOrigin.Begin);

            return File(npoiMemoryStream, "application/stream", $"模板流文件无头部底部{Guid.NewGuid()}.xlsx");
        }

        /// <summary>
        /// 根据模板文件流导出数据到excel 
        /// 进一步选择列
        /// </summary>
        /// <param name="templateFormFile">模板文件</param>
        /// <returns></returns>
        [HttpPost]
        public async Task<FileResult> ExportObjectToExcelByTemplateSelect(IFormFile templateFormFile)
        {
            var streamone = templateFormFile.OpenReadStream();

            var streamTwo = new MemoryStream();
            streamone.CopyTo(streamTwo);
            streamone.Seek(0, SeekOrigin.Begin);
            streamTwo.Seek(0, SeekOrigin.Begin);

            List<string> colNames = new List<string>();
            colNames.Add("姓名");
            colNames.Add("出生日期");

            var streamThree = new MemoryStream();
            streamone.CopyTo(streamThree);
            streamThree.Seek(0, SeekOrigin.Begin);

            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(2);
            ExcelHandle.getExcelTemplate().setDataStartColNum(2);




            ExcelHandle.exportObjectToExcelByTemplate(streamTwo, $@"D:\excelexportstreamtemplate\指定路径模板{Guid.NewGuid().ToString("N")}.xlsx", null, excelDataOutDtos, colNames);



            //导出到流
            ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(2);
            ExcelHandle.getExcelTemplate().setDataStartColNum(2);
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();
            ExcelHandle.exportObjectToExcelByTemplate(streamThree, npoiMemoryStream, null, excelDataOutDtos, colNames);
            npoiMemoryStream.Seek(0, SeekOrigin.Begin);

            return File(npoiMemoryStream, "application/stream", $"模板流文件无头部底部{Guid.NewGuid()}.xlsx");
        }

        /// <summary>
        /// 模板文件中存在列名的导出
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public async Task<FileResult> ExportObjectToExcelHasColNameByStream(IFormFile templateFormFile)
        {
            var streamone = templateFormFile.OpenReadStream();

            var streamTwo = new MemoryStream();
            streamone.CopyTo(streamTwo);
            streamone.Seek(0, SeekOrigin.Begin);
            streamTwo.Seek(0, SeekOrigin.Begin);
            var streamThree = new MemoryStream();
            streamone.CopyTo(streamThree);
            streamThree.Seek(0, SeekOrigin.Begin);
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);
            ExcelHandle.setIsClasspath(false);
            ExcelHandle.getExcelTemplate().setIsOrder(true);
            ExcelHandle.getExcelTemplate().insertOrder();
            ExcelHandle.getExcelTemplate().setOrderName("排序号");//给序号列设置名称
            Dictionary<string, object> headerAndFooterMaps = new Dictionary<string, object>();
            headerAndFooterMaps.Add("datetime", DateTime.Now);
            headerAndFooterMaps.Add("exporter", "OpenDeepSpace管理员");
            headerAndFooterMaps.Add("count", excelDataOutDtos.Count);
            ExcelHandle.exportObjectToExcelByTemplateHasColName(streamTwo, $@"D:\excelexporthftemplate\指定路径带头部底部模板{Guid.NewGuid().ToString("N")}.xlsx", headerAndFooterMaps, excelDataOutDtos);

            ExcelHandle = new ExcelHandle();
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);
            ExcelHandle.getExcelTemplate().insertOrder();
            ExcelHandle.getExcelTemplate().setOrderName("排序号");
            ExcelHandle.exportObjectToExcelByTemplateHasColName(streamThree, npoiMemoryStream, headerAndFooterMaps, excelDataOutDtos);
            npoiMemoryStream.Seek(0, SeekOrigin.Begin);


            return File(npoiMemoryStream, "application/stream", $"模板流文件有头部底部{Guid.NewGuid()}.xlsx");

        }

        /// <summary>
        /// 模板文件中存在列名的导出
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public async Task<FileResult> ExportObjectToExcelHasColName()
        {
            ExcelHandle ExcelHandle = new ExcelHandle();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);
            ExcelHandle.setIsClasspath(false);
            ExcelHandle.getExcelTemplate().setIsOrder(true);
            ExcelHandle.getExcelTemplate().setOrderName("排序号");//给序号列设置名称
            Dictionary<string, object> headerAndFooterMaps = new Dictionary<string, object>();
            headerAndFooterMaps.Add("datetime", DateTime.Now);
            headerAndFooterMaps.Add("exporter", "OpenDeepSpace管理员");
            headerAndFooterMaps.Add("count", excelDataOutDtos.Count);
            ExcelHandle.exportObjectToExcelByTemplateHasColName(@"C:\Users\lenovo\Desktop\导出的数据模板含头部底部有列名.xlsx", $@"D:\excelexporthftemplate\指定路径带头部底部模板{Guid.NewGuid().ToString("N")}.xlsx", headerAndFooterMaps, excelDataOutDtos);

            ExcelHandle = new ExcelHandle();
            NpoiMemoryStream npoiMemoryStream = new NpoiMemoryStream();
            ExcelHandle.getExcelTemplate().setExcelColNums(4);
            ExcelHandle.getExcelTemplate().insertOrder();
            ExcelHandle.getExcelTemplate().setOrderName("排序号");
            ExcelHandle.exportObjectToExcelByTemplateHasColName(@"C:\Users\lenovo\Desktop\导出的数据模板含头部底部有列名.xlsx", npoiMemoryStream, headerAndFooterMaps, excelDataOutDtos);
            npoiMemoryStream.Seek(0, SeekOrigin.Begin);


            return File(npoiMemoryStream, "application/stream", $"模板流文件有头部底部{Guid.NewGuid()}.xlsx");

        }

        /// <summary>
        /// 导出excel数据到对象
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public void ExportExcelToObject()
        {
            ExcelHandle ExcelHandle = new ExcelHandle();
            var excelDatas = ExcelHandle.exportExcelToObject<ExcelDataOutDto>(@"C:\Users\lenovo\Desktop\Excel数据导出到对象.xlsx");
        }

        /// <summary>
        /// 导出excel数据到对象 
        /// 注意数据列和Col特性中的列名要对应
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public void ExportExcelToObjectByStream(IFormFile formFile)
        {
            ExcelHandle ExcelHandle = new ExcelHandle();
            var excelDatas = ExcelHandle.exportExcelToObject<ExcelDataOutDto>(formFile.OpenReadStream());
        }

    }
}
