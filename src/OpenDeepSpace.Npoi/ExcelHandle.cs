using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using OpenDeepSpace.Npoi.Attributes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace OpenDeepSpace.Npoi
{
    /// <summary>
    /// Excel处理
    /// </summary>
    public class ExcelHandle
    {
		
		/// <summary>
		/// 是否模板是在classpath下面 默认在classpath路径下查找 即是否放在工程入口某个目录下
		/// </summary>
		private bool isClasspath = true;


		/// <summary>
		/// 是否是输出2007及以上版本的xlsx文件默认为true
		/// </summary>
		private bool isXSSF = true;



		/// <summary>
		/// 是否设置序号针对无模板，默认设置
		/// </summary>
		private bool isSetOrder = true;

		public bool IsSetOrder()
		{
			return isSetOrder;
		}

		/// <summary>
		/// 设置是否显示序列号
		/// </summary>
		/// <param name="isSetOrder"></param>
		public void setIsSetOrder(bool isSetOrder)
		{
			this.isSetOrder = isSetOrder;
			this.excelTemplate.setIsOrder(isSetOrder);
		}

		public bool IsXSSF()
		{
			return isXSSF;
		}

		/// <summary>
		/// 设置为false后表示  是以2003版本的xls文件
		/// </summary>
		/// <param name="isXSSF"></param>
		public void setIsXSSF(bool isXSSF)
		{
			this.isXSSF = isXSSF;
		}

		/// <summary>
		/// excel模板
		/// </summary>
		private ExcelTemplate excelTemplate = new ExcelTemplate();
		/// <summary>
		/// 日期格式对象
		/// </summary>
		private ExcelDatePattern excelDatePattern = new ExcelDatePattern();

		public ExcelDatePattern getExcelDatePattern()
		{ 
			return excelDatePattern;
		}

		public ExcelTemplate getExcelTemplate()
		{
			return excelTemplate;
		}


		public void setIsClasspath(bool isClasspath)
		{
			this.isClasspath = isClasspath;
		}

		public ExcelHandle()
		{
			excelTemplate.setExcelDatePattern(this.excelDatePattern);

		}


		/// <summary>
		/// 处理对象到Excel
		/// </summary>
		/// <param name="stream">模板流</param>
		/// <param name="infos">头部及底部类似的其他信息</param>
		/// <param name="objs">对象集合</param>
		/// <param name="colNames">所需导出列的名字的集合 外部传入</param>
		/// <param name="isSetColName"></param>
		private void handleObjectToExcel<T>(Stream stream, Dictionary<string, object> infos, List<T> objs,
				List<string> colNames, bool isSetColName)
		{

			excelTemplate.readExcelTemplateByStream(stream);

			setExcelColName(typeof(T), colNames, isSetColName);
			setDatas(objs, colNames);


			if (infos != null)
			{

				getExcelTemplate().replaceExcelTemplate(infos);
			}

		}

		/// <summary>
		/// 处理对象到Excel
		/// </summary>
		/// <param name="excelTemplatePath">模板路径</param>
		/// <param name="infos">头部及底部类似的其他信息</param>
		/// <param name="objs">对象集合</param>
		/// <param name="colNames">所需导出列的名字的集合 外部传入</param>
		/// <param name="isSetColName"></param>
		private void handleObjectToExcel<T>(string excelTemplatePath, Dictionary<string, object> infos, List<T> objs,
				 List<string> colNames, bool isSetColName)
		{

			if (isClasspath)
			{

				excelTemplate.readExcelTemplateByClasspath(excelTemplatePath);

			}
			else
			{

				excelTemplate.readExcelTemplateByFilePath(excelTemplatePath);
			}

			setExcelColName(typeof(T), colNames, isSetColName);
			setDatas(objs,  colNames);


			if (infos != null)
			{

				getExcelTemplate().replaceExcelTemplate(infos);
			}

		}


		/// <summary>
		/// 处理多个Sheet表格
		/// </summary>
		/// <param name="excelTemplatePath"></param>
		/// <param name="exportExcelPath"></param>
		/// <param name="infosList"></param>
		/// <param name="objsList"></param>
		/// <param name="sheetNum"></param>
		public void exportObjectToExcelMoreSheetByTemplate<T>(string excelTemplatePath, string exportExcelPath, List<Dictionary<string, object>> infosList, List<List<T>> objsList, int sheetNum)
		{

			handleObjectToExcelMoreSheet(excelTemplatePath, infosList, objsList, null, false, sheetNum);


			excelTemplate.writeDataToExcel(exportExcelPath);
		}

		/// <summary>
		/// 处理多个Sheet表格
		/// </summary>
		/// <param name="excelTemplatePath"></param>
		/// <param name="exportExcelStream"></param>
		/// <param name="infosList"></param>
		/// <param name="objsList"></param>
		/// <param name="sheetNum"></param>
		public void exportObjectToExcelMoreSheetByTemplate<T>(string excelTemplatePath, Stream exportExcelStream, List<Dictionary<string, object>> infosList, List<List<T>> objsList, int sheetNum)
		{

			handleObjectToExcelMoreSheet(excelTemplatePath, infosList, objsList, null, false, sheetNum);


			excelTemplate.writeDataToStream(exportExcelStream);
		}


		/// <summary>
		/// 处理多个Sheet表格
		/// </summary>
		/// <param name="stream"></param>
		/// <param name="exportExcelPath"></param>
		/// <param name="infosList"></param>
		/// <param name="objsList"></param>
		/// <param name="sheetNum"></param>
		public void exportObjectToExcelMoreSheetByTemplate<T>(Stream stream, string exportExcelPath, List<Dictionary<string, object>> infosList, List<List<T>> objsList, int sheetNum)
		{

			handleObjectToExcelMoreSheet(stream, infosList, objsList, null, false, sheetNum);


			excelTemplate.writeDataToExcel(exportExcelPath);
		}

		/// <summary>
		/// 处理多个Sheet表格
		/// </summary>
		/// <param name="stream">模板流</param>
		/// <param name="exportExcelStream">导出的excel流</param>
		/// <param name="infosList"></param>
		/// <param name="objsList"></param>
		/// <param name="sheetNum"></param>
		public void exportObjectToExcelMoreSheetByTemplate<T>(Stream stream, Stream exportExcelStream, List<Dictionary<string, object>> infosList, List<List<T>> objsList,int sheetNum)
		{

			handleObjectToExcelMoreSheet(stream, infosList, objsList,  null, false, sheetNum);


			excelTemplate.writeDataToStream(exportExcelStream);
		}


		/// <summary>
		/// 多个工作表处理
		/// </summary>
		/// <param name="excelTemplatePath"></param>
		/// <param name="infosList"></param>
		/// <param name="objsList"></param>
		/// <param name="colNames"></param>
		/// <param name="isSetColName"></param>
		/// <param name="sheetNum"></param>
		private void handleObjectToExcelMoreSheet<T>(string excelTemplatePath, List<Dictionary<string, object>> infosList, List<List<T>> objsList,
				 List<string> colNames, bool isSetColName, int sheetNum)
		{
			if (isClasspath)
			{

				excelTemplate.readExcelTemplateByClasspath(excelTemplatePath, sheetNum);

			}
			else
			{

				excelTemplate.readExcelTemplateByFilePath(excelTemplatePath, sheetNum);
			}

			setExcelColName(typeof(T), colNames, isSetColName);
			setDatas(objsList[0], colNames);
			if (infosList != null && infosList.Count > 0)
			{

				getExcelTemplate().replaceExcelTemplate(infosList[0]);
			}

			if (sheetNum > 1)
			{
				for (int sheetIndex = 0; sheetIndex < sheetNum - 1; sheetIndex++)
				{

					excelTemplate.initExcelTemplate(sheetIndex + 1);//在初始化一个sheet
																	//又开始设置数据
					setExcelColName(typeof(T), colNames, isSetColName);

					if (sheetIndex + 1 < objsList.Count)
						setDatas(objsList[sheetIndex + 1],colNames);
					if (infosList != null && infosList.Count > 0 && sheetIndex + 1 < infosList.Count)
					{

						getExcelTemplate().replaceExcelTemplate(infosList[sheetIndex + 1]);
					}
				}
			}



		}

		/// <summary>
		/// 多个工作表处理
		/// </summary>
		/// <param name="stream"></param>
		/// <param name="infosList"></param>
		/// <param name="objsList"></param>
		/// <param name="colNames"></param>
		/// <param name="isSetColName"></param>
		/// <param name="sheetNum"></param>
		private void handleObjectToExcelMoreSheet<T>(Stream stream, List<Dictionary<string, object>> infosList, List<List<T>> objsList,
				 List<string> colNames, bool isSetColName, int sheetNum)
		{
			
			excelTemplate.readExcelTemplateByStream(stream, sheetNum);

			setExcelColName(typeof(T), colNames, isSetColName);
			setDatas(objsList[0], colNames);
			if (infosList != null && infosList.Count > 0)
			{

				getExcelTemplate().replaceExcelTemplate(infosList[0]);
			}

			if (sheetNum > 1)
			{
				for (int sheetIndex = 0; sheetIndex < sheetNum - 1; sheetIndex++)
				{

					excelTemplate.initExcelTemplate(sheetIndex + 1);//在初始化一个sheet
																	//又开始设置数据
					setExcelColName(typeof(T), colNames, isSetColName);

					if (sheetIndex + 1 < objsList.Count)
						setDatas(objsList[sheetIndex + 1], colNames);
					if (infosList != null && infosList.Count > 0 && sheetIndex + 1 < infosList.Count)
					{

						getExcelTemplate().replaceExcelTemplate(infosList[sheetIndex + 1]);
					}
				}
			}



		}


		/// <summary>
		/// 导出对象到Excel 基于某一个Excel模板 带头部底部以及类似的信息
		/// </summary>
		/// <param name="stream">模板的路径</param>
		/// <param name="exportExcelPath">根据模板导出的Excel的路径</param>
		/// <param name="infos">头部以及底部类似信息的集合</param>
		/// <param name="objs">对象集合</param>
		public void exportObjectToExcelByTemplate<T>(Stream stream, string exportExcelPath, Dictionary<string, object> infos, List<T> objs)
		{

			handleObjectToExcel(stream, infos, objs,null, true);


			excelTemplate.writeDataToExcel(exportExcelPath);
		}

		/// <summary>
		/// 导出对象到Excel 基于某一个Excel模板 带头部底部以及类似的信息
		/// </summary>
		/// <param name="excelTemplatePath">模板的路径</param>
		/// <param name="exportExcelStream">导出的excel流</param>
		/// <param name="infos">头部以及底部类似信息的集合</param>
		/// <param name="objs">对象集合</param>
		public void exportObjectToExcelByTemplate<T>(string excelTemplatePath, Stream exportExcelStream, Dictionary<string, object> infos, List<T> objs)
		{

			handleObjectToExcel(excelTemplatePath, infos, objs, null, true);


			excelTemplate.writeDataToStream(exportExcelStream);
		}

		/// <summary>
		/// 导出对象到Excel 基于某一个Excel模板 带头部底部以及类似的信息
		/// </summary>
		/// <param name="stream">模板的路径</param>
		/// <param name="exportExcelStream">根据模板导出的Excel的路径</param>
		/// <param name="infos">头部以及底部类似信息的集合</param>
		/// <param name="objs">对象集合</param>
		public void exportObjectToExcelByTemplate<T>(Stream stream, Stream exportExcelStream, Dictionary<string, object> infos, List<T> objs)
		{

			handleObjectToExcel(stream, infos, objs,null, true);


			excelTemplate.writeDataToStream(exportExcelStream);
		}

		/// <summary>
		/// 导出对象到Excel 基于某一个Excel模板 带头部底部以及类似的信息
		/// </summary>
		/// <param name="excelTemplatePath">模板的路径</param>
		/// <param name="exportExcelPath">根据模板导出的Excel的路径</param>
		/// <param name="infos">头部以及底部类似信息的集合</param>
		/// <param name="objs">对象集合</param>
		public void exportObjectToExcelByTemplate<T>(string excelTemplatePath, string exportExcelPath, Dictionary<string, object> infos, List<T> objs)
		{

			handleObjectToExcel(excelTemplatePath, infos, objs,  null, true);


			excelTemplate.writeDataToExcel(exportExcelPath);
		}



		/// <summary>
		/// 导出对象到Excel 基于某一个Excel模板  无头部底部以及类似的信息
		/// </summary>
		/// <param name="excelTemplatePath">模板的路径</param>
		/// <param name="exportExcelPath">根据模板导出的Excel的路径</param>
		/// <param name="objs">对象集合</param>
		public void exportObjectToExcelByTemplate<T>(string excelTemplatePath, string exportExcelPath, List<T> objs)
		{

			exportObjectToExcelByTemplate(excelTemplatePath, exportExcelPath, null, objs);
		}

		/// <summary>
		/// 导出对象到Excel 基于某一个Excel模板  无头部底部以及类似的信息
		/// </summary>
		/// <param name="excelTemplatePath">模板的路径</param>
		/// <param name="exportExcelStream">根据模板导出Excel流</param>
		/// <param name="objs">对象集合</param>
		public void exportObjectToExcelByTemplate<T>(string excelTemplatePath, Stream exportExcelStream, List<T> objs)
		{

			exportObjectToExcelByTemplate(excelTemplatePath, exportExcelStream, null, objs);
		}

		/// <summary>
		/// 导出对象到Excel 基于某一个Excel模板  无头部底部以及类似的信息
		/// </summary>
		/// <param name="stream">模板流</param>
		/// <param name="exportExcelPath">根据模板导出的Excel的路径</param>
		/// <param name="objs">对象集合</param>
		public void exportObjectToExcelByTemplate<T>(Stream stream, string exportExcelPath, List<T> objs)
		{

			exportObjectToExcelByTemplate(stream, exportExcelPath, null, objs);
		}

		/// <summary>
		/// 导出对象到Excel 基于某一个Excel模板  无头部底部以及类似的信息
		/// </summary>
		/// <param name="stream">模板流</param>
		/// <param name="exportExcelStream">根据模板导出的Excel流</param>
		/// <param name="objs">对象集合</param>
		public void exportObjectToExcelByTemplate<T>(Stream stream, Stream exportExcelStream, List<T> objs)
		{

			exportObjectToExcelByTemplate(stream, exportExcelStream, null, objs);
		}


		/// <summary>
		/// 带列名的模板导出
		/// </summary>
		/// <param name="excelTemplatePath">模板路径</param>
		/// <param name="exportExcelPath">导出的excel路径</param>
		/// <param name="infos">头部底部信息</param>
		/// <param name="objs">数据集合</param>
		public void exportObjectToExcelByTemplateHasColName<T>(string excelTemplatePath, string exportExcelPath, Dictionary<string, object> infos, List<T> objs)
		{

			handleObjectToExcel(excelTemplatePath, infos, objs,null, false);

			excelTemplate.writeDataToExcel(exportExcelPath);
		}

		/// <summary>
		/// 带列名的模板导出
		/// </summary>
		/// <param name="excelTemplatePath">模板路径</param>
		/// <param name="exportExcelStream">导出的excel路径</param>
		/// <param name="infos">头部底部信息</param>
		/// <param name="objs">数据集合</param>
		public void exportObjectToExcelByTemplateHasColName<T>(string excelTemplatePath, Stream exportExcelStream, Dictionary<string, object> infos, List<T> objs)
		{

			handleObjectToExcel(excelTemplatePath, infos, objs,  null, false);

			excelTemplate.writeDataToStream(exportExcelStream);
		}

		/// <summary>
		/// 带列名的模板导出
		/// </summary>
		/// <param name="excelTemplateStream">模板文件流</param>
		/// <param name="exportExcelPath">导出的excel路径</param>
		/// <param name="infos">头部底部信息</param>
		/// <param name="objs">数据集合</param>
		public void exportObjectToExcelByTemplateHasColName<T>(Stream excelTemplateStream, string exportExcelPath, Dictionary<string, object> infos, List<T> objs)
		{

			handleObjectToExcel(excelTemplateStream, infos, objs,null, false);

			excelTemplate.writeDataToExcel(exportExcelPath);
		}

		/// <summary>
		/// 带列名的模板导出
		/// </summary>
		/// <param name="excelTemplateStream">模板流文件</param>
		/// <param name="exportExcelStream">导出的excel模板流</param>
		/// <param name="infos">头部底部信息</param>
		/// <param name="objs">数据集合</param>
		public void exportObjectToExcelByTemplateHasColName<T>(Stream excelTemplateStream, Stream exportExcelStream, Dictionary<string, object> infos, List<T> objs)
		{

			handleObjectToExcel(excelTemplateStream, infos, objs,  null, false);

			excelTemplate.writeDataToStream(exportExcelStream);
		}


		/// <summary>
		/// 导出对象到Excel基于某一个模板 可以选择字段导出所需字段内容
		/// </summary>
		/// <param name="excelTemplatePath">模板路径</param>
		/// <param name="exportExcelPath">导出的excel路径</param>
		/// <param name="infos">头部底部信息</param>
		/// <param name="objs">数据对象集合</param>
		/// <param name="colNames">需要导出的字段</param>
		public void exportObjectToExcelByTemplate<T>(string excelTemplatePath, string exportExcelPath, Dictionary<string, object> infos, List<T> objs, List<string> colNames)
		{

			handleObjectToExcel(excelTemplatePath, infos, objs, colNames, true);


			excelTemplate.writeDataToExcel(exportExcelPath);
		}

		/// <summary>
		/// 导出对象到Excel基于某一个模板 可以选择字段导出所需字段内容
		/// </summary>
		/// <param name="excelTemplateStream">模板流</param>
		/// <param name="exportExcelPath">导出的excel存储路径</param>
		/// <param name="infos">头部底部信息</param>
		/// <param name="objs">数据对象集合</param>
		/// <param name="colNames">需要导出的字段</param>
		public void exportObjectToExcelByTemplate<T>(Stream excelTemplateStream, string exportExcelPath, Dictionary<string, object> infos, List<T> objs, List<string> colNames)
		{

			handleObjectToExcel(excelTemplateStream, infos, objs, colNames, true);


			excelTemplate.writeDataToExcel(exportExcelPath);
		}

		/// <summary>
		/// 导出对象到Excel基于某一个模板 可以选择字段导出所需字段内容
		/// </summary>
		/// <param name="excelTemplatePath">模板路径</param>
		/// <param name="exportExcelStream">导出的excel流</param>
		/// <param name="infos">头部底部信息</param>
		/// <param name="objs">对象数据集合</param>
		/// <param name="colNames">需要导出的字段</param>
		public void exportObjectToExcelByTemplate<T>(string excelTemplatePath, Stream exportExcelStream, Dictionary<string, object> infos, List<T> objs, List<string> colNames)
		{

			handleObjectToExcel(excelTemplatePath, infos, objs, colNames, true);


			excelTemplate.writeDataToStream(exportExcelStream);
		}

		/// <summary>
		/// 导出对象到Excel基于某一个模板 可以选择字段导出所需字段内容
		/// </summary>
		/// <param name="excelTemplateStream">excel模板流</param>
		/// <param name="exportExcelStream">导出的excel流</param>
		/// <param name="infos">头部底部信息</param>
		/// <param name="objs">对象数据集合</param>
		/// <param name="colNames">需要导出的字段</param>
		public void exportObjectToExcelByTemplate<T>(Stream excelTemplateStream, Stream exportExcelStream, Dictionary<string, object> infos, List<T> objs, List<string> colNames)
		{

			handleObjectToExcel(excelTemplateStream, infos, objs, colNames, true);


			excelTemplate.writeDataToStream(exportExcelStream);
		}


		/// <summary>
		/// 无Excel模板的数据写入到Excel中
		/// </summary>
		/// <param name="exportExcelPath">导出Excel的路径</param>
		/// <param name="objs">对象集合</param>
		/// <exception cref="NpoiException"></exception>
		public void exportObjectToExcel<T>(string exportExcelPath, List<T> objs)
		{

			//根据后缀名智能判断是否是2007以上的excel
			if (!string.Equals(System.IO.Path.GetExtension(exportExcelPath), ".xlsx"))
				setIsXSSF(false);

			IWorkbook wb = handleObjectToExcel(objs,  null);


			BufferedStream bs = null;
			FileStream fs = null;

			//判断文件夹是否存在 不存在则创建
			string path=System.IO.Path.GetDirectoryName(exportExcelPath);
			if(!System.IO.Directory.Exists(path))
				System.IO.Directory.CreateDirectory(path);

			using (fs = new FileStream(exportExcelPath, FileMode.Create))
			{
				try
				{
					bs = new BufferedStream(fs);
					wb.Write(bs);

				}
				catch (IOException e)
				{
					// TODO Auto-generated catch block

					throw new NpoiException("写入对象到不带模板的Excel文件失败");
				}
				finally
				{

					try
					{

						if (bs != null)
						{
							bs.Close();
						}

					}
					catch (Exception e)
					{


					}
				}


			}

		}

		/// <summary>
		/// 无Excel模板的数据写入到Excel 通过colMaps指定导出列
		/// </summary>
		/// <param name="exportExcelPath">导出excel的路径</param>
		/// <param name="objs">对象集合</param>
		/// <param name="colNames">指定导出列以及给列进行重命名</param>
		/// <exception cref="NpoiException"></exception>
		public void exportObjectToExcel<T>(string exportExcelPath, List<T> objs, List<string> colNames)
		{
			//根据后缀名智能判断是否是2007以上的excel
			if (!string.Equals(System.IO.Path.GetExtension(exportExcelPath), ".xlsx"))
				setIsXSSF(false);
			IWorkbook wb = handleObjectToExcel(objs, colNames);


			BufferedStream bs = null;
			FileStream fs = null;

			//判断文件夹是否存在 不存在则创建
			string path = System.IO.Path.GetDirectoryName(exportExcelPath);
			if (!System.IO.Directory.Exists(path))
				System.IO.Directory.CreateDirectory(path);

			using (fs = new FileStream(exportExcelPath, FileMode.Create))
			{
				try
				{
					bs = new BufferedStream(fs);
					wb.Write(bs);

				}
				catch (IOException e)
				{
					// TODO Auto-generated catch block

					throw new NpoiException("写入对象到不带模板的Excel文件失败");
				}
				finally
				{

					try
					{

						if (bs != null)
						{
							bs.Close();
						}

					}
					catch (Exception e)
					{


					}
				}


			}

		}

		/// <summary>
		/// 无Excel模板的数据写入Excel并到流
		/// </summary>
		/// <param name="os">流</param>
		/// <param name="objs">对象数据集合</param>
		/// <exception cref="NpoiException"></exception>
		public void exportObjectToExcel<T>(Stream os, List<T> objs)
		{

			IWorkbook wb = handleObjectToExcel(objs, null);

			BufferedStream bos = null;
			try
			{
				bos = new BufferedStream(os);
				wb.Write(bos);

			}
			catch (IOException e)
			{


				throw new NpoiException("写入对象到不带模板的流中失败");
			}
			finally
			{

				try
				{

					if (bos != null)
					{
						bos.Close();
					}
					wb.Close();

				}
				catch (Exception e)
				{


				}
			}

		}

		/// <summary>
		/// 无Excel模板的数据写入到Excel并到流
		/// </summary>
		/// <param name="os">流</param>
		/// <param name="objs">数据对象集合</param>
		/// <param name="colNames">指定列名导出</param>
		/// <exception cref="NpoiException"></exception>
		public void exportObjectToExcel<T>(Stream os, List<T> objs, List<string> colNames)
		{

			IWorkbook wb = handleObjectToExcel(objs, colNames);

			BufferedStream bos = null;
			try
			{
				bos = new BufferedStream(os);
				wb.Write(bos);

			}
			catch (IOException e)
			{


				throw new NpoiException("写入对象到不带模板的流中失败");
			}
			finally
			{

				try
				{

					if (bos != null)
					{
						bos.Close();
					}
					wb.Close();

				}
				catch (Exception e)
				{


				}
			}

		}

		/// <summary>
		/// 处理无模板数据到Excel的转换
		/// </summary>
		/// <param name="objs"></param>
		/// <param name="colNames">可选列名集合</param>
		/// <returns></returns>
		private IWorkbook handleObjectToExcel<T>(List<T> objs, List<string> colNames)
        {

            IWorkbook wb = null;

            if (isXSSF)
            {

                wb = new XSSFWorkbook();

            }
            else
            {

                wb = new HSSFWorkbook();
            }

            //考虑这里循环导出多个sheet 数据超过65535 考虑以65535分割或者提示异常
            List<dynamic> list = new List<dynamic>();
            foreach (T obj in objs)//数据转换为dynamic
            {
                list.Add(obj);
            }

            //获取是否存在使用ExcelSheet标注的属性
            RecursionFill(list, colNames, wb, typeof(T));

            return wb;
        }


        /// <summary>
        /// 导出json数据到excel
        /// 形如：[{"ColumnTitles":["姓名","部门"],"SheetName":"第一个sheet名称","Datas":[{"Name":"小李","Department":"管理部"}],"IsOrder":true},{}]
		/// 形如：[{"Datas":[{"Name":"小李","Department":"管理部"}]},{}]
        /// </summary>
        /// <param name="os"></param>
        /// <param name="JsonData"></param>
        public void exportJsonToExcel(Stream os, string JsonData)
		{ 
			JArray jArray=Newtonsoft.Json.JsonConvert.DeserializeObject<JArray>(JsonData);

            IWorkbook wb = null;

            if (isXSSF)
            {

                wb = new XSSFWorkbook();

            }
            else
            {

                wb = new HSSFWorkbook();
            }

            foreach (JObject item in jArray)
			{

				ISheet sheet = null;
                //ExcelColumn名称
                List<ExcelColumn> excelColumns = new List<ExcelColumn>();

				//检查是否包含IsOrder
				if (item.ContainsKey("IsOrder"))
				{//是否排序 
				   isSetOrder = (bool)item["IsOrder"];
				}

				//检查是否包含SheetName属性
				if (item.ContainsKey("SheetName"))
				{//对应sheetName 
					string SheetName = item["SheetName"].ToString();
					sheet = wb.CreateSheet(SheetName);
				}
				else
				{
					sheet = wb.CreateSheet();
				}


				//检查是否包含ColumnTitles属性
				if (item.ContainsKey("ColumnTitles"))
				{ //对应ExcelColumn

					if (item["ColumnTitles"] is JArray)
					{
					    JArray columnTitles = item["ColumnTitles"] as JArray;

                        int i = 1;
						foreach (var columnTitle in columnTitles)
                        {
                            ExcelColumn excelColumn = new ExcelColumn();
                            excelColumn.ColOrder = i;
                            excelColumn.ColName = columnTitle.ToString();
                            excelColumns.Add(excelColumn);
                            i++;
                        }

						//设置列名
                        SetExcelColumnName(sheet, excelColumns);


                    }
					
				}


				//检查是否包含Datas属性
				if (item.ContainsKey("Datas"))
				{//对应动态数据

					//解析数据 数据应该是集合形式的JArray
					if (item["Datas"] is JArray)
					{
						var datas = item["Datas"] as JArray;

						List<Dictionary<int,object>> list= new List<Dictionary<int, object>>();
						foreach (JObject data in datas)
						{

							//是否已经设置了ColumnTitles
							bool IsSetColumnTitles = false;

							//取属性
							List<JProperty> jproperties = data.Properties().ToList();

                            //如果不包含ColumnTitles  则以数据中的属性名称为列标题 且未设置
                            if (!item.ContainsKey("ColumnTitles") && !IsSetColumnTitles)
							{
								var jpropertyNames = jproperties.Select(t => t.Name.ToString()).ToList();
								
								int i = 1;
								foreach (string colname in jpropertyNames)
								{


									excelColumns.Add(new ExcelColumn()
									{
										ColOrder = i,
										ColName= colname,
									});
									i++;
								}

								SetExcelColumnName(sheet, excelColumns);

							}

							//遍历每一数据属性及其值
							Dictionary<int,object> dicData= new Dictionary<int, object>();
							int index = 1;
							foreach (JProperty jproperty in jproperties)
							{
								dicData.Add(index, jproperty.Value.ToString());
								index++;

							}
							list.Add(dicData);

						}

						//设置数据
						setDatas(list, sheet, excelColumns);

					
					}

				}


			}


            BufferedStream bos = null;
            try
            {
                bos = new BufferedStream(os);
                wb.Write(bos);

            }
            catch (IOException e)
            {


                throw new NpoiException("写入对象到不带模板的流中失败");
            }
            finally
            {

                try
                {

                    if (bos != null)
                    {
                        bos.Close();
                    }
                    wb.Close();

                }
                catch (Exception e)
                {


                }
            }
        }


        private void SetExcelColumnName(ISheet sheet, List<ExcelColumn> excelColumns)
        {
            //创建列标题
            //创建列名所在行
            IRow row = sheet.CreateRow(0);

            int startColIndex = 0;

            if (isSetOrder)
            {
                row.CreateCell(0).SetCellValue(excelTemplate.getOrderName());
            }

            foreach (ExcelColumn ec in excelColumns)
            {

                if (isSetOrder)
                    startColIndex++;

                //输出列名
                row.CreateCell(startColIndex).SetCellValue(ec.ColName);

                if (!isSetOrder) startColIndex++;
            }
        }

        /// <summary>
        /// 不带模板的设置数据到单元格
        /// </summary>
        /// <param name="objs">对象集合</param>
        /// <param name="sheet">单元格</param>
		/// <param name="excelColumns">excel列名</param>
        private void setDatas(List<Dictionary<int,object>> objs, ISheet sheet, List<ExcelColumn> excelColumns)
        {
            IRow row = null;

            

            //写数据
            for (int i = 0; i < objs.Count; i++)
            {

                row = FillRowData(sheet, excelColumns, i, objs[i]);

            }


        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="objs"></param>
        /// <param name="colNames"></param>
        /// <param name="wb"></param>
        /// <param name="dataType"></param>
        private void RecursionFill(List<dynamic> objs, List<string> colNames, IWorkbook wb, Type dataType)
        {

			//自身填充
			FillSheet(colNames, wb, objs, dataType);

            var includeSheets = dataType.GetProperties().Where(p => p.GetCustomAttribute<ExcelSheetAttribute>() != null).ToList();
            foreach (var includeSheet in includeSheets)
            {
                //获取属性包含的数据类型
                Type singleDataType = GetDataType(includeSheet);

                //获取关联数据
                List<dynamic> dynamics = new List<dynamic>();//动态数据
                foreach (var item in objs)
                {
                    dynamic relationDatas = includeSheet.GetValue(item, null);
                    if (relationDatas != null)
                        dynamics.AddRange(relationDatas);
                }

                RecursionFill(dynamics, colNames, wb, singleDataType);


            }
        }

        /// <summary>
        /// 填充sheet
        /// </summary>
        /// <param name="colNames"></param>
        /// <param name="wb"></param>
        /// <param name="list"></param>
        /// <param name="dataType"></param>
        private void FillSheet(List<string> colNames, IWorkbook wb, List<dynamic> list, Type dataType)
        {
            ISheet sheet = null;
            //获取是否存在ExcelSheet特性
            ExcelSheetAttribute excelSheetAttribute = dataType.GetCustomAttribute<ExcelSheetAttribute>();

            //不为空并且SheetName不为空
            if (excelSheetAttribute != null && !string.IsNullOrWhiteSpace(excelSheetAttribute.SheetName))
            {
                sheet = wb.CreateSheet(excelSheetAttribute.SheetName);
            }
            else
            {
                sheet = wb.CreateSheet();

            }

            //TODO: ColNames考虑键值对 或顺序设置 对不同关联数据导出的字段进行筛选
            setExcelColName(sheet, dataType, colNames);


            setDatas(list, dataType, sheet, colNames);

            //合并数据列
            MergeDataCol(list, dataType, sheet, colNames);
        }

        /// <summary>
        /// 获取数据类型
        /// </summary>
        /// <param name="includeSheet"></param>
        /// <returns></returns>
        private static Type GetDataType(PropertyInfo includeSheet)
        {
            Type singleDataType = includeSheet.PropertyType;//数据的类型
                                                            //获取实际类型
            if (singleDataType.IsGenericType)//如果是泛型
            {
                singleDataType = singleDataType.GetGenericArguments()[0];
            }
            else//如果是数组
            {

            }

            return singleDataType;
        }

        /// <summary>
        /// 不带模板的设置数据到单元格
        /// </summary>
        /// <param name="objs">对象集合</param>
		/// <param name="dataType">数据类型</param>
        /// <param name="sheet">单元格</param>
        /// <param name="colNames">可选的列名集合</param>
        private void setDatas(List<dynamic> objs,Type dataType, ISheet sheet, List<string> colNames)
        {
            IRow row = null;

            List<ExcelColumn> excelColumns = GetDataExcelColumn(colNames, dataType);
            excelColumns.Sort();//排序

            //写数据
            for (int i = 0; i < objs.Count; i++)
            {

                row = FillRowData(sheet, excelColumns, i, objs[i], dataType);
                
            }


        }

        /// <summary>
        /// 填充行数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="excelColumns"></param>
        /// <param name="i"></param>
		/// <param name="dataDic"></param>
        /// <returns></returns>
        private IRow FillRowData(ISheet sheet, List<ExcelColumn> excelColumns, int i, Dictionary<int,object> dataDic)
        {
            //创建一行
            IRow row = sheet.CreateRow(i + 1);
            if (isSetOrder)
            {//设置了序号列，每次每行的第一列设置序号

                row.CreateCell(0).SetCellValue(i + 1);
            }

            for (int index = 0; index < excelColumns.Count; index++)
            {
                //创建列 并填充相应的数据
                ICell cell = null;
                if (isSetOrder)
                    cell = row.CreateCell(index + 1);
                else
                    cell = row.CreateCell(index);


                object data = dataDic[index+1];//取数据

                dataInputCell(cell, data.GetType(), data, excelColumns[index]);
            }

            return row;
        }


        private List<ExcelColumn> GetDataExcelColumn(List<string> colNames,Type dataType)
        {
            List<ExcelColumn> excelColumns = null;
            if (colNames != null && colNames.Count > 0)
            {
                excelColumns = getExcelColumns(dataType, colNames);
            }
            else
            {
                excelColumns = getExcelColumns(dataType);
            }

            return excelColumns;
        }

        /// <summary>
        /// 填充行数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="excelColumns"></param>
        /// <param name="i"></param>
        /// <param name="obj"></param>
        /// <param name="dataType"></param>
        /// <returns></returns>
        private IRow FillRowData(ISheet sheet, List<ExcelColumn> excelColumns, int i, object obj, Type dataType)
        {
            //创建一行
            IRow row = sheet.CreateRow(i + 1);
            if (isSetOrder)
            {//设置了序号列，每次每行的第一列设置序号

                row.CreateCell(0).SetCellValue(i + 1);
            }

            for (int index = 0; index < excelColumns.Count; index++)
            {
                //创建列 并填充相应的数据
                ICell cell = null;
                if (isSetOrder)
                    cell = row.CreateCell(index + 1);
                else
                    cell = row.CreateCell(index);


                object data = dataType.GetProperty(excelColumns[index].PropertyName).GetValue(obj);

                DataValidation(excelColumns[index], data);

                dataInputCell(cell, excelColumns[index].PropertyType, data, excelColumns[index]);
            }

            return row;
        }

		/// <summary>
		/// 合并数据列
		/// </summary>
		/// <param name="objs"></param>
		/// <param name="dataType">数据类型</param>
		/// <param name="sheet"></param>
		/// <param name="colNames"></param>
		private void MergeDataCol(List<dynamic> objs,Type dataType, ISheet sheet, List<string> colNames)
		{

			List<ExcelColumn> excelColumns = GetDataExcelColumn(colNames, dataType);
			excelColumns.Sort();//排序

			//合并行列
			//合并的序号开始下标
			int orderIndex = 1;
			//是否需要合并列 合并列就统计相同的行数
			//筛选出需要合并的列
			excelColumns = excelColumns.Where(t => t.MergeCol).ToList();
			//找出基准列
			ExcelColumn baseLineExcelColumn = excelColumns.FirstOrDefault(t => t.IsBaselineCol);
			if (baseLineExcelColumn != null)
			{//存在基准列
				int baseLineColStartRow = 1, baseLineColEndRow;//基准列开始行和结束行
				if (isSetOrder == false)//不设置下标 colOrder-1
					baseLineExcelColumn.ColOrder--;
				//去掉基准列
				excelColumns = excelColumns.Where(t => !t.IsBaselineCol).ToList();

				//设置当前值 取第一个值为当前值
				object currentValue = dataType.GetProperty(baseLineExcelColumn.PropertyName).GetValue(objs[0]);

				for (int j = 0; j < objs.Count; j++)
				{

					//根据反射获取数据
					object tempValue = dataType.GetProperty(baseLineExcelColumn.PropertyName).GetValue(objs[j]);

					if (currentValue != tempValue || (currentValue == tempValue && j == objs.Count - 1))//如果当前值和获取出来的值不相同表示合并开始新的数据行 或者全部数据遍历完成
					{
						baseLineColEndRow = j + 1;//之前的结束列
												  //如果存在大于一行的数据
						if (baseLineColEndRow - baseLineColStartRow > 1)
						{ //存在需要合并的数据才合并

							if (j != objs.Count - 1)//非最后一行 合并时的结束列需要减1
							{
								baseLineColEndRow--;
							}
							//合并序号列并重新赋值 如果存在
							if (isSetOrder)
							{
                                //AddMergedRegion效率低 使用AddMergedRegionUnsafe替换
                                sheet.AddMergedRegionUnsafe(new CellRangeAddress(baseLineColStartRow, baseLineColEndRow, 0, 0));
								sheet.GetRow(baseLineColStartRow).GetCell(0).SetCellValue(orderIndex);
							}
							//合并之前列 //合并基准列
							sheet.AddMergedRegionUnsafe(new CellRangeAddress(baseLineColStartRow, baseLineColEndRow, baseLineExcelColumn.ColOrder, baseLineExcelColumn.ColOrder));
							//合并其他列
							foreach (ExcelColumn otherExcelColumn in excelColumns)
							{
								sheet.AddMergedRegionUnsafe(new CellRangeAddress(baseLineColStartRow, baseLineColEndRow, otherExcelColumn.ColOrder, otherExcelColumn.ColOrder));
							}
						}
						else 
						if (baseLineColEndRow - baseLineColStartRow == 1)////如果只有一行差距 不需要合并
						{
							sheet.GetRow(baseLineColStartRow).GetCell(0).SetCellValue(orderIndex);

							//如果已是最后一行数据
							if (j == objs.Count - 1)//重置下标
							{
								sheet.GetRow(baseLineColStartRow+1).GetCell(0).SetCellValue(orderIndex+1);
							}
						}
						
						

						//开始新的列
						baseLineColStartRow = j + 1;//由于存在列名关系所以开始行+1

						//当前值附新值
						currentValue = tempValue;

						//不同的值的时候下标滚动
						orderIndex++;//下标滚动
					}
				}



			}


			//是否需要要合并行 合并行就统计相同的列数 
		}


		/// <summary>
		/// 数据填充到不到Excel模板的文件中
		/// </summary>
		/// <param name="cell">单元格</param>
		/// <param name="typeClass"></param>
		/// <param name="data">数据</param>
		/// <param name="excelColumn"></param>
		private void dataInputCell(ICell cell, Type typeClass, object data,ExcelColumn excelColumn=null)
		{

			if (data == null || data.ToString().Equals(""))
			{
				cell.SetCellValue("");

			}
			else
			{


				//根据返回值类型设置数据
				if (typeClass == typeof(int))
				{

					cell.SetCellValue(int.Parse(data.ToString()));

				}
				else if (typeClass == typeof(bool))
				{

					cell.SetCellValue(bool.Parse(data.ToString()));
				}
				else if (typeClass == typeof(DateTime))
				{
					SimpleDateFormat sdf = new SimpleDateFormat(excelDatePattern.getDatePattern());
					if(!string.IsNullOrWhiteSpace(excelColumn.DatePattern))//如果ExcelColumn中存在时间格式化
						sdf=new SimpleDateFormat(excelColumn.DatePattern);

					cell.SetCellValue(sdf.Format((DateTime)data));

				}
				else if (typeClass == typeof(double))
				{

					cell.SetCellValue(double.Parse(data.ToString()));

				}/*else if(typeClass==typeof(Calendar)){
				
				cell.SetCellValue((Calendar) data);
				
			}*/
				else if (typeClass == typeof(string))
				{

					cell.SetCellValue(data.ToString());
				}


			}

		}

		/// <summary>
		/// 设置数据 有Excel模板的完成对象到Excel的转换
		/// </summary>
		/// <param name="objs"></param>
		/// <param name="colNames">指定列名</param>
		private void setDatas<T>(List<T> objs, List<string> colNames)
		{


			List<ExcelColumn> excelColumns = null;
			if (colNames != null && colNames.Count > 0)
			{

				excelColumns = getExcelColumns(typeof(T), colNames);
			}
			else
			{
				excelColumns = getExcelColumns(typeof(T));
			}
			excelColumns.Sort();//排序


			foreach (object obj in objs)
			{//一个对象对应一行数据

				//TODO 
				foreach (ExcelColumn ec in excelColumns)
                {


                    object data = typeof(T).GetProperty(ec.PropertyName).GetValue(obj);

                    DataValidation(ec,data);

                    //TODO 这个方法里面需要优化
                    dataInputCell(ec.PropertyType, data);
                }
            }

		}

		/// <summary>
		/// 数据验证
		/// </summary>
		/// <param name="ec"></param>
		/// <exception cref="NpoiException"></exception>
		private static void DataValidation(ExcelColumn ec, object data,int rowIndex)
		{
			//数据验证
			var dataValidationAttrs = ec.PropertyInfo.GetCustomAttributes().Where(t => t is DataValidationAttribute);

			foreach (var dataValidationAttr in dataValidationAttrs)
			{
				//调用数据验证
				var dataValidationResult = (dataValidationAttr as DataValidationAttribute).IsValid(data);
				if (dataValidationResult != null)
					throw new NpoiException(NpoiExceptionCode.DataException,rowIndex==-1?dataValidationResult.ErrorMessage:dataValidationResult.ErrorMessage.Replace("{row}",$"{rowIndex}"));
			}
			
        }

		/// <summary>
		/// 数据验证
		/// </summary>
		/// <param name="ec"></param>
		/// <exception cref="NpoiException"></exception>
		private static void DataValidation(ExcelColumn ec, object data)
		{
			DataValidation(ec, data, -1);

		}


		/// <summary>
		/// 数据填充到带Excel模板的输出的Excel文件中
		/// </summary>
		/// <param name="typeClass"></param>
		/// <param name="data"></param>
		private void dataInputCell(Type typeClass, object data)
		{


			//需要设置序号 插入序号
			if (excelTemplate.IsOrder()) excelTemplate.insertOrder(1, excelTemplate.getInsertOrderRowIndex() + 1);

			if (data == null || data.Equals(""))
			{

				excelTemplate.inputRowCell("");

			}
			else
			{


				//根据返回值类型设置数据
				if (typeClass == typeof(int))
				{
					excelTemplate.inputRowCell(int.Parse(data.ToString()));
				}
				else if (typeClass == typeof(bool))
				{

					excelTemplate.inputRowCell(bool.Parse(data.ToString()));
				}
				else if (typeClass == typeof(DateTime))
				{

					excelTemplate.inputRowCell((DateTime)data);
				}
				else if (typeClass == typeof(double))
				{

					excelTemplate.inputRowCell(double.Parse(data.ToString()));
				}/*else if(typeClass==typeof(Calendar)){
				
				excelTemplate.inputRowCell((Calendar) data);
			}*/
				else if (typeClass == typeof(string))
				{

					excelTemplate.inputRowCell(data.ToString());
				}

			}
		}


		/// <summary>
		/// 为带模板的Excel设置列名
		/// </summary>
		/// <param name="type"></param>
		/// <param name="colNames">导出列名的集合</param>
		private void setExcelColName(Type type, List<string> colNames)
		{
			setExcelColName(type, colNames, true);
		}

		/// <summary>
		/// 带模板的设置列名（是否）
		/// </summary>
		/// <param name="type"></param>
		/// <param name="colNames"></param>
		/// <param name="isSetColName"></param>
		private void setExcelColName(Type type, List<string> colNames, bool isSetColName)
		{
			List<ExcelColumn> excelColumns = null;
			if (colNames != null && colNames.Count > 0)
			{
				excelColumns = getExcelColumns(type, colNames);
			}
			else
			{

				excelColumns = getExcelColumns(type);
			}
			excelColumns.Sort();//排序
			if (excelTemplate.IsOrder())
			{
				excelTemplate.setOrderIndentifier();
			}
			if (isSetColName == false) return;

			foreach (ExcelColumn ec in excelColumns)
			{

				//输出列名
				excelTemplate.inputRowCell(ec.ColName);
			}
		}

		/// <summary>
		/// 为不使用模板的Excel设置列名
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="type"></param>
		/// <param name="colNames">列名集合</param>
		private void setExcelColName(ISheet sheet, Type type, List<string> colNames)
		{
			List<ExcelColumn> excelColumns = null;
			if (colNames != null && colNames.Count > 0)
			{
				excelColumns = getExcelColumns(type,colNames);
			}
			else
			{
				excelColumns = getExcelColumns(type);
			}
			excelColumns.Sort();//排序
								//创建列名所在行
			IRow row = sheet.CreateRow(0);

			int startColIndex = 0;

			if (isSetOrder)
			{
				row.CreateCell(0).SetCellValue(excelTemplate.getOrderName());
			}

			foreach (ExcelColumn ec in excelColumns)
			{

				if (isSetOrder)
					startColIndex++;

				//输出列名
				row.CreateCell(startColIndex).SetCellValue(ec.ColName);

				if (!isSetOrder) startColIndex++;
			}
		}

		/// <summary>
		/// 获取一个ExcelColumn类的集合
		/// </summary>
		/// <param name="type">类型</param>
		/// <param name="colNames">指定列名集合</param>
		/// <returns></returns>
		private List<ExcelColumn> getExcelColumns(Type type, List<string> colNames)
		{

			List<ExcelColumn> excelColumns = new List<ExcelColumn>();

			//读取包含ExcelColumnAttribute特性的属性
			PropertyInfo[] propertyInfos = type.GetProperties();
			//判断是否有多个基准列 如果存在多个提示错误
			if (propertyInfos.Select(t => t.GetCustomAttribute<ExcelColumnAttribute>()).Where(t =>t!=null && t.IsBaselineCol).Count()>1)
				throw new NpoiException($"{type.FullName}存在超过一个基准列");

			foreach (PropertyInfo property in propertyInfos)
			{


				//哪些属性上面用了ExcelAttribute特性的
				if (Attribute.IsDefined(property, typeof(ExcelColumnAttribute)))
				{
					

					//获取ExcelColumnAttribute标注的内容
					ExcelColumnAttribute excelColumnAttribute = (ExcelColumnAttribute)Attribute.GetCustomAttribute(property, typeof(ExcelColumnAttribute));
					
					ExcelColumn excelColumn = null;
					//maps中存在的元素 就添加到集合中
					if (colNames != null && colNames.Count > 0)
					{
						//循环比对
						foreach (string colName in colNames)
						{
							
							if (excelColumnAttribute.ColName.Trim().Equals(colName.Trim()))
							{

								excelColumn = new ExcelColumn(excelColumnAttribute.ColName, excelColumnAttribute.ColOrder, property.Name, property.PropertyType);
								excelColumn.MergeRow = excelColumnAttribute.MergeRow;
								excelColumn.MergeCol = excelColumnAttribute.MergeColumn;
								excelColumn.IsBaselineCol = excelColumnAttribute.IsBaselineCol;
								excelColumn.IsCalculateCol = excelColumnAttribute.IsCalculateCol;
								excelColumn.DatePattern = excelColumnAttribute.DatePattern;
								excelColumn.PropertyInfo = property;
								excelColumns.Add(excelColumn);
							}

						}
					}
					else
					{
						excelColumn = new ExcelColumn(excelColumnAttribute.ColName, excelColumnAttribute.ColOrder, property.Name, property.PropertyType);
						excelColumn.MergeRow = excelColumnAttribute.MergeRow;
						excelColumn.MergeCol = excelColumnAttribute.MergeColumn;
						excelColumn.IsBaselineCol = excelColumnAttribute.IsBaselineCol;
						excelColumn.IsCalculateCol = excelColumnAttribute.IsCalculateCol;
						excelColumn.DatePattern= excelColumnAttribute.DatePattern;
						excelColumn.PropertyInfo= property;
						excelColumns.Add(excelColumn);
					}

				}

			}
			return excelColumns;
		}

		/// <summary>
		/// 获取一个ExcelColumn集合 实体中所有标记了ExcelColumnAttribute特性
		/// </summary>
		/// <param name="type"></param>
		/// <returns></returns>
		private List<ExcelColumn> getExcelColumns(Type type)
		{

			List<ExcelColumn> excelColumns = getExcelColumns(type, null);

			return excelColumns;
		}

		/// <summary>
		/// 导出无头部与底部的Excel中的数据到对象
		/// </summary>
		/// <param name="excelPath"></param>
		/// <param name="sheetIndex">工作表下标</param>
		/// <returns></returns>
		public List<T> exportExcelToObject<T>(string excelPath, int sheetIndex)
		{


			List<T> objs = exportExcelToObject<T>(excelPath,  0, 0, sheetIndex);


			return objs;
		}

		/// <summary>
		/// 导出无头部与底部的Excel中的数据到对象
		/// </summary>
		/// <param name="excelStream"></param>
		/// <param name="sheetIndex">工作表下标</param>
		/// <param name = "IsAutoSkipNullRow" > 是否自动跳过空行 默认<see cref="true"/> 即跳过</param>
		/// <returns></returns>
		public List<T> exportExcelToObject<T>(Stream excelStream, int sheetIndex,bool IsAutoSkipNullRow=true)
		{


			List<T> objs = exportExcelToObject<T>(excelStream, 0, 0, sheetIndex,IsAutoSkipNullRow);


			return objs;
		}


		/// <summary>
		/// 导出无头部与底部的Excel中的数据到对象
		/// 只有一个sheet
		/// </summary>
		/// <param name="excelPath"></param>
		/// <returns></returns>
		public List<T> exportExcelToObject<T>(string excelPath)
		{

			List<T> objs = exportExcelToObject<T>(excelPath,  0);

			return objs;
		}

		/// <summary>
		/// 导出无头部与底部的Excel中的数据到对象
		/// 只有一个sheet
		/// </summary>
		/// <param name="excelStream"></param>
		/// <param name = "IsAutoSkipNullRow" > 是否自动跳过空行 默认<see cref="true"/> 即跳过</param>
		/// <returns></returns>
		public List<T> exportExcelToObject<T>(Stream excelStream,bool IsAutoSkipNullRow=true)
		{

			List<T> objs = exportExcelToObject<T>(excelStream, 0,IsAutoSkipNullRow);

			return objs;
        }



		/// <summary>
		/// 导出Excel的数据到对象
		/// </summary>
		/// <param name="excelPath"></param>
		/// <param name="colNameStartRow"></param>
		/// <param name="notDataRowNum">如果获取的数据由问题可适当调整一下该值为负数 不是数据的行数 这里涉及到总行数减去非数据行数 这样来算出数据结束行</param>
		/// <param name="sheetIndex"></param>
		/// <param name = "IsAutoSkipNullRow" > 是否自动跳过空行 默认<see cref="true"/> 即跳过</param>
		/// <returns></returns>
		public List<T> exportExcelToObject<T>(string excelPath, int colNameStartRow, int notDataRowNum, int sheetIndex,bool IsAutoSkipNullRow=true)
		{

			List<T> objs = new List<T>();

			handleExcelToObject(excelPath,colNameStartRow, notDataRowNum, sheetIndex, objs,IsAutoSkipNullRow);


			return objs;
		}

        /// <summary>
        /// 导出Excel的数据到对象
        /// </summary>
        /// <param name="excelStream"></param>
        /// <param name="colNameStartRow"></param>
        /// <param name="notDataRowNum">如果获取的数据由问题可适当调整一下该值为负数</param>
        /// <param name="sheetIndex"></param>
        /// <param name="IsAutoSkipNullRow">是否自动跳过空行 默认<see cref="true"/>即跳过</param>
        /// <returns></returns>
        public List<T> exportExcelToObject<T>(Stream excelStream,  int colNameStartRow, int notDataRowNum, int sheetIndex,bool IsAutoSkipNullRow=true)
		{

			List<T> objs = new List<T>();

			handleExcelToObject(excelStream,colNameStartRow, notDataRowNum, sheetIndex, objs,IsAutoSkipNullRow);


			return objs;
		}

		/// <summary>
		/// 导出有头部有底部只有一个工作表的Excel
		/// </summary>
		/// <param name="excelPath"></param>
		/// <param name="colNameStartRow"></param>
		/// <param name="notDataRowNum"></param>
		/// <returns></returns>
		public List<T> exportExcelToObject<T>(string excelPath, int colNameStartRow, int notDataRowNum)
		{


			List<T> objs = exportExcelToObject<T>(excelPath,  colNameStartRow, notDataRowNum, 0);

			return objs;
		}

		/// <summary>
		/// 导出有头部有底部只有一个工作表的Excel
		/// </summary>
		/// <param name="excelStream"></param>
		/// <param name="colNameStartRow"></param>
		/// <param name="notDataRowNum"></param>
		/// <returns></returns>
		public List<T> exportExcelToObject<T>(Stream excelStream, int colNameStartRow, int notDataRowNum)
		{


			List<T> objs = exportExcelToObject<T>(excelStream, colNameStartRow, notDataRowNum, 0);

			return objs;
		}



		/// <summary>
		/// 处理Excel的数据到对象
		/// </summary>
		/// <param name="excelPath">excel的文件路径</param>
		/// <param name="colNameStartRow">列标识所在行</param>
		/// <param name="notDataRowNum">没有数据的行</param>
		/// <param name="sheetIndex">工作表的下标</param>
		/// <param name="objs">对象集合</param>
		/// <param name="IsAutoSkipNullRow">是否自动跳过空行 默认<see cref="true"/>即跳过</param>
		/// <exception cref="NpoiException"></exception>
		private void handleExcelToObject<T>(string excelPath,int colNameStartRow, int notDataRowNum,
				int sheetIndex, List<T> objs,bool IsAutoSkipNullRow=true)
		{

			IWorkbook wb = null;

			FileStream fs = null;
			try
			{
				using (fs = new FileStream(excelPath, FileMode.Open))
				{

					if (isClasspath)
					{//从classpath读取

						wb = WorkbookFactory.Create(fs);

					}
					else
					{

						wb = WorkbookFactory.Create(fs);

					}

					ISheet sheet = wb.GetSheetAt(sheetIndex);

					Dictionary<int, ExcelColumn> colPropertyMaps = getColProperty(sheet.GetRow(colNameStartRow), typeof(T));


					readExcelData(colNameStartRow, notDataRowNum, objs, sheet, colPropertyMaps,IsAutoSkipNullRow);


				}


			}
			catch (InvalidFormatException e)
			{
				// TODO Auto-generated catch block
				e.ToString();
			}
			catch (IOException e)
			{
				if (fs != null) fs.Close();
				throw new NpoiException("读取Excel失败");
			}
			catch (SecurityException e)
			{
				// TODO Auto-generated catch block
				e.ToString();
			}
		}

        /// <summary>
        /// 处理Excel的数据到对象
        /// </summary>
        /// <param name="colNameStartRow">列标识所在行</param>
        /// <param name="notDataRowNum">没有数据的行</param>
        /// <param name="sheetIndex">工作表的下标</param>
        /// <param name="objs">对象集合</param>
        /// <param name="IsAutoSkipNullRow">是否自动跳过空行 默认<see cref="true"/>即跳过</param>
        /// <exception cref="NpoiException"></exception>
        private void handleExcelToObject<T>(Stream excelStream,int colNameStartRow, int notDataRowNum,
				int sheetIndex, List<T> objs, bool IsAutoSkipNullRow = true)
		{

			IWorkbook wb = null;

			try
			{
				
					
					wb = WorkbookFactory.Create(excelStream);

					

					ISheet sheet = wb.GetSheetAt(sheetIndex);


					Dictionary<int, ExcelColumn> colPropertyMaps = getColProperty(sheet.GetRow(colNameStartRow), typeof(T));

					readExcelData(colNameStartRow, notDataRowNum, objs, sheet, colPropertyMaps,IsAutoSkipNullRow);




			}
			catch (InvalidFormatException e)
			{
				// TODO Auto-generated catch block
				e.ToString();
			}
			catch (IOException e)
			{
				
				throw new NpoiException("读取Excel失败");
			}
			catch (SecurityException e)
			{
				// TODO Auto-generated catch block
				e.ToString();
			}
		}

		/// <summary>
		/// 读取Excel中的数据
		/// </summary>
		/// <param name="colNameStartRow"></param>
		/// <param name="notDataRowNum"></param>
		/// <param name="objs"></param>
		/// <param name="sheet"></param>
		/// <param name="IsAutoSkipNullRow"></param>
		/// <param name="colPropertyMaps"></param>
		private void readExcelData<T>( int colNameStartRow, int notDataRowNum, List<T> objs, ISheet sheet,
				Dictionary<int, ExcelColumn> colPropertyMaps,bool IsAutoSkipNullRow=true)
		{

			//这还有一点小问题 就是行数的判断 暂未解决 可以通过noDataRowNum来调整
			for (int rowIndex = colNameStartRow + 1; rowIndex <= sheet.LastRowNum - notDataRowNum; rowIndex++)
			{
				//读取数据
				T obj = (T)Activator.CreateInstance(typeof(T), true);

				IRow dataRow = sheet.GetRow(rowIndex);

				if (dataRow == null)//行数据为空直接跳过
				{
					if (!IsAutoSkipNullRow)
					{
						//不是自动跳过 判断第一个数据验证
						DataValidation(colPropertyMaps[0], typeof(T).GetProperty(colPropertyMaps[0].PropertyName).GetValue(obj), rowIndex);
					}

					continue;
				
				}

				//一行空列计数
				int nullColumn = 0;//默认为0

				foreach (int colIndex in colPropertyMaps.Keys)
				{

					ICell dataCell = dataRow.GetCell(colIndex);
					ExcelColumn excelColumnProperty;
					string colIndexPropertyName;
					colPropertyMaps.TryGetValue(colIndex, out excelColumnProperty);
					colIndexPropertyName = excelColumnProperty.PropertyName;
					Type paramTypeClass;
					//记录方法的传入参数的类型
					Dictionary<string, Type> propertyParamTypeMaps = new Dictionary<string, Type>();
					foreach (PropertyInfo propertyinfo in typeof(T).GetProperties())
					{

						propertyParamTypeMaps.Add(propertyinfo.Name, propertyinfo.PropertyType);


					}
					propertyParamTypeMaps.TryGetValue(colIndexPropertyName, out paramTypeClass);


					//空cell
					if (dataCell == null)//自动跳过 
					{ 
						nullColumn++;
						//设置为自动跳过
						if (!IsAutoSkipNullRow)
						{ //不自动跳过就数据验证
							DataValidation(excelColumnProperty,typeof(T).GetProperty(colIndexPropertyName).GetValue(obj),rowIndex);
						}
						continue;//跳过
					}

					CellType cellType = dataCell.CellType;

					if (cellType == CellType.Blank)//Blank如何处理 有格式这些比如单元格填充色 但数据内容为空 ：跳过 还是 继续判断验证
					{
						nullColumn++;
						
					}




					PropertyInfo propertyInfo = null;
					if ((paramTypeClass == typeof(int)) && cellType == CellType.Numeric)
					{

						propertyInfo = typeof(T).GetProperty(colIndexPropertyName);

						var data = Convert.ToInt32(dataCell.NumericCellValue);
						DataValidation(excelColumnProperty, data,rowIndex);

						propertyInfo.SetValue(obj, Convert.ToInt32(data));
					}
					else if (paramTypeClass == typeof(double))
					{

						propertyInfo = typeof(T).GetProperty(colIndexPropertyName);

						var data = dataCell.NumericCellValue;
						DataValidation(excelColumnProperty, data,rowIndex);

						propertyInfo.SetValue(obj, data);

					}
					else if (paramTypeClass == typeof(string))
					{

						propertyInfo = typeof(T).GetProperty(colIndexPropertyName);

						//Cannot get a text value from a numeric cell
						dataCell.SetCellType(CellType.String);

						var data = dataCell.StringCellValue;
						DataValidation(excelColumnProperty, data,rowIndex);

						propertyInfo.SetValue(obj, data);

					}
					else if (paramTypeClass == typeof(DateTime))
					{

						propertyInfo = typeof(T).GetProperty(colIndexPropertyName);
						SimpleDateFormat sdf = new SimpleDateFormat(excelDatePattern.getDatePattern());
						DateTime date = new DateTime();
						try
						{
							if (!dataCell.StringCellValue.Trim().Equals(""))
								date = sdf.Parse(dataCell.StringCellValue.Trim());
						}
						catch 
						{
							date = dataCell.DateCellValue;
						}
						finally
						{

							DataValidation(excelColumnProperty, date,rowIndex);

						}

						propertyInfo.SetValue(obj, date);

					}



				}

				if (IsAutoSkipNullRow && colPropertyMaps.Keys.Count() == nullColumn)//自动跳过空行
					continue;

				objs.Add(obj);

			}
		}


		/// <summary>
		/// 完成列下标和方法的对应
		/// </summary>
		/// <param name="colIndentifierRow"></param>
		/// <param name="type"></param>
		/// <returns></returns>
		private Dictionary<int, ExcelColumn> getColProperty(IRow colIndentifierRow, Type type)
		{

			//存储列下标以及对应的方法名
			Dictionary<int, ExcelColumn> maps = new Dictionary<int, ExcelColumn>();

			List<ExcelColumn> ecColumns = getExcelColumns(type);
			ecColumns.Sort();

			if (colIndentifierRow == null)
				throw new NpoiException(NpoiExceptionCode.ColumnNameException,"列标识行为空");

			foreach (ICell cell in colIndentifierRow)
			{//遍历excel文件中的标题行

				string cellColName = cell.StringCellValue.Trim();

				foreach (ExcelColumn ec in ecColumns)
				{//Excel中的列标识比对

					if (ec.ColName.Trim().Equals(cellColName))
					{
						maps.Add(cell.ColumnIndex, ec);
						break;
					}


				}
			}

			//如果excel中的列名 缺少
			if (maps.Count != ecColumns.Count)
			{ 
				var loseColumns = ecColumns.Where(t=>maps.Values.All(a=>a!=t));

				throw new NpoiException(NpoiExceptionCode.ColumnNameException,$"Excel文件中缺少列名:{string.Join(",",loseColumns.Select(t=>t.ColName))}");
			}

			if(maps.Count==0)
				throw new NpoiException(NpoiExceptionCode.ColumnNameException,"读取列标识错误，请检查列标识开始行的位置");

			return maps;
		}

	}
}
