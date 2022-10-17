using NPOI;
using NPOI.OpenXml4Net.Exceptions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDeepSpace.Npoi
{
	/// <summary>
	/// ExcelTemplate
	/// </summary>
	public class ExcelTemplate
	{

		private IWorkbook workbook = null;

		/// <summary>
		/// 日期格式对象
		/// </summary>
		private ExcelDatePattern excelDatePattern = new ExcelDatePattern();


		public void setExcelDatePattern(ExcelDatePattern excelDatePattern)
		{ 
			this.excelDatePattern = excelDatePattern;
		}

		public ExcelDatePattern getExcelDatePattern()
		{
			return this.excelDatePattern;
		}

		/// <summary>
		/// 对应模板中从哪一个单元格开始填数据,填充数据的标识符
		/// </summary>
		private static string DATACELL_INDENTIFIER = "datas";

		/// <summary>
		/// 样式标识，对应模板中的styles
		/// </summary>
		private static string STYLES_INDENTIFIER = "styles";

		/// <summary>
		/// 序号标识符，对应模板中的order
		/// </summary>
		private static string ORDER_INDENTIFIER = "order";

		/// <summary>
		/// 记录需要插入序号的列下标
		/// </summary>
		private int insertOrderColIndex = -1;
		/// <summary>
		/// 记录需要插入序号的行下标
		/// </summary>
		private int insertOrderRowIndex = -1;

		/// <summary>
		/// 默认为模板中没找到order
		/// </summary>
		private bool ifFindOrder = false;

		/// <summary>
		/// 获取插入序号的行下标
		/// </summary>
		/// <returns></returns>
		public int getInsertOrderRowIndex()
		{
			return insertOrderRowIndex;
		}

		/// <summary>
		/// 是否设置列号,默认为true
		/// </summary>
		private bool isOrder = true;

		/// <summary>
		/// 存储样式标识对应的列号以及相应的样式
		/// </summary>
		Dictionary<int, ICellStyle> styleMaps = new Dictionary<int, ICellStyle>();

		/// <summary>
		/// 开始填充数据的单元格所在行
		/// </summary>
		private int startDataCellRowIndex = -1;
		/// <summary>
		/// 开始填充数据的单元格所在列
		/// </summary>
		private int startDataCellColIndex = -1;

		/// <summary>
		/// 要填充数据的单元格当前行
		/// </summary>
		private int currentDataCellRowIndex;

		/// <summary>
		/// 当前数据行
		/// </summary>
		/// <returns></returns>
		public int getCurrentDataCellRowIndex()
		{
			return currentDataCellRowIndex;
		}

		/// <summary>
		/// 要填充数据的单元格当前列
		/// </summary>
		private int currentDataCellColIndex;

		/// <summary>
		/// 工作表
		/// </summary>
		private ISheet sheet;

		/// <summary>
		/// 当前填数据的行
		/// </summary>
		private IRow currentRow;

		/// <summary>
		/// sheet的下标
		/// </summary>
		private int sheetIndex = 0;

		/// <summary>
		/// 记录最后一行的位置
		/// </summary>
		private int lastRowIndex;

		/// <summary>
		/// 默认样式
		/// </summary>
		private ICellStyle defaultCellStyle;
		/// <summary>
		/// 标识默认样式，对应模板中的defaultStyles
		/// </summary>
		private static string DEFAULT_STYLE_INDENTIFIER = "defaultStyles";

		/// <summary>
		/// 默认行高
		/// </summary>
		private float defaultRowHeight;


		/// <summary>
		/// 数据类型枚举
		/// </summary>
		public enum TYPE
		{

			STring, DOuble, INteger, DAte, BOolean, CAlendar
		}

		/// <summary>
		/// excel模板的列个数
		/// </summary>
		private int excelColNums = 1;//默认存在一个列即序号列

		/// <summary>
		/// 单元格数据为空占位符
		/// </summary>
		private string nullPlaceHolder = "";

		/// <summary>
		/// 序号
		/// </summary>
		private string orderName = "序号";


		public string getOrderName()
		{
			return orderName;
		}

		public void setOrderName(string orderName)
		{
			this.orderName = orderName;
		}


		public bool IsOrder()
		{
			return isOrder;
		}

		/// <summary>
		/// 设置是否显示序列号
		/// </summary>
		/// <param name="isOrder"></param>
		public void setIsOrder(bool isOrder)
		{
			this.isOrder = isOrder;
		}

		public int getSheetIndex()
		{
			return sheetIndex;
		}

		public void setSheetIndex(int sheetIndex)
		{
			this.sheetIndex = sheetIndex;
		}

		public ExcelTemplate()
		{

		}




		public int getExcelColNums()
		{
			return excelColNums;
		}

		/// <summary>
		/// 设置excel列
		/// </summary>
		/// <param name="excelColNums"></param>
		public void setExcelColNums(int excelColNums)
		{
			this.excelColNums += excelColNums;

		}

		/// <summary>
		/// 设置数据开始列 从0开始计算
		/// </summary>
		/// <param name="dataStartColNum"></param>
		public void setDataStartColNum(int dataStartColNum)
		{
			this.excelColNums += dataStartColNum;
		}

		
		FileStream fs = null;

		/// <summary>
		/// 从工程的classpath下面读取Excel模板
		/// </summary>
		/// <param name="path"></param>
		/// <returns></returns>
		public ExcelTemplate readExcelTemplateByClasspath(string path)
		{

			readExcelTemplateByClasspath(path, 1);
			return this;
		}

		/// <summary>
		/// 根据一个sheet模板需要写出多个sheet
		/// </summary>
		/// <param name="path"></param>
		/// <param name="sheetNum"></param>
		/// <returns></returns>
		/// <exception cref="NpoiException"></exception>
		public ExcelTemplate readExcelTemplateByClasspath(string path, int sheetNum)
		{

			if (excelColNums == -1)
			{

				throw new NpoiException("需要设置excel的列数");
			}

			try
			{

				using (fs = new FileStream(path, FileMode.Open, FileAccess.Read))
				{
					workbook = WorkbookFactory.Create(fs);



					if (sheetNum > 1)
					{ //至少是导出两个sheet

						//根据模板sheet复制一个sheet(这里workbook克隆sheet并不会改变本身模板内容)
						for (int i = 0; i < sheetNum - 1; i++)
						{

							workbook.CloneSheet(0);

						}

					}
					/*else {//只有一个sheet（在克隆模板会改变xlsx文件的情况下）
						//如果存在多个sheet,删除多个sheet只保留一个
						if (workbook.NumberOfSheets > 1) {
							for (int i = 1; i < workbook.NumberOfSheets; i++) {
								workbook.RemoveSheetAt(i);
							}
						}
					
					}*/

					initExcelTemplate();
				}

			}
			catch (EncryptedDocumentException e)
			{

				throw new NpoiException("读取Excel模板格式有误!");

			}
			catch (IOException e)
			{

				if (fs != null)
				{
					fs.Close();
				}
				return this;
				throw new NpoiException("读取Excel模板不存在或正在被打开不可用IOException");

			}
			catch (InvalidFormatException e)
			{

				throw new NpoiException(e.StackTrace);
			}


			return this;
		}

		/// <summary>
		/// 读取外部文件的模板
		/// </summary>
		/// <param name="path"></param>
		/// <returns></returns>
		public ExcelTemplate readExcelTemplateByFilePath(string path)
		{
			readExcelTemplateByFilePath(path, 1);

			return this;
		}

		public ExcelTemplate readExcelTemplateByFilePath(string path, int sheetNum)
		{
			if (excelColNums == -1)
			{

				throw new NpoiException("需要设置excel的列数");
			}
			try
			{
				workbook = WorkbookFactory.Create(File.OpenRead(path));
				if (sheetNum > 1)
				{ //至少是导出两个sheet

					//根据模板sheet复制一个sheet(这里workbook克隆sheet并不会改变本身模板内容)
					for (int i = 0; i < sheetNum - 1; i++)
					{

						workbook.CloneSheet(0);

					}

				}
				/*else {//只有一个sheet（在克隆模板会改变xlsx文件的情况下）
					//如果存在多个sheet,删除多个sheet只保留一个
					if (workbook.NumberOfSheets > 1) {
						for (int i = 1; i < workbook.NumberOfSheets; i++) {
							workbook.RemoveSheetAt(i);
						}
					}
				}*/
				initExcelTemplate();
			}
			catch (EncryptedDocumentException e)
			{

				throw new NpoiException("读取Excel模板格式有误!");

			}
			catch (IOException e)
			{

				throw new NpoiException("读取Excel模板不存在或正在其他程序打开");
			}
			catch (InvalidFormatException e)
			{
				
				throw new NpoiException(e.StackTrace);
			}

			return this;
		}

		/// <summary>
		/// 从流读取文件的模板
		/// </summary>
		/// <param name="stream"></param>
		/// <returns></returns>
		public ExcelTemplate readExcelTemplateByStream(Stream stream)
		{
			readExcelTemplateByStream(stream, 1);

			return this;
		}

		public ExcelTemplate readExcelTemplateByStream(Stream stream, int sheetNum)
		{
			if (excelColNums == -1)
			{

				throw new NpoiException("需要设置excel的列数");
			}
			try
			{
				workbook = WorkbookFactory.Create(stream);
				if (sheetNum > 1)
				{ //至少是导出两个sheet

					//根据模板sheet复制一个sheet(这里workbook克隆sheet并不会改变本身模板内容)
					for (int i = 0; i < sheetNum - 1; i++)
					{

						workbook.CloneSheet(0);

					}

				}
				/*else {//只有一个sheet（在克隆模板会改变xlsx文件的情况下）
					//如果存在多个sheet,删除多个sheet只保留一个
					if (workbook.NumberOfSheets > 1) {
						for (int i = 1; i < workbook.NumberOfSheets; i++) {
							workbook.RemoveSheetAt(i);
						}
					}
				}*/
				initExcelTemplate();
			}
			catch (EncryptedDocumentException e)
			{

				throw new NpoiException("读取Excel模板格式有误!");

			}
			catch (IOException e)
			{

				throw new NpoiException("读取Excel模板不存在或正在其他程序打开");
			}
			catch (InvalidFormatException e)
			{

				throw new NpoiException(e.StackTrace);
			}

			return this;
		}



		/// <summary>
		/// 初始化模板获取sheet工作表
		/// </summary>
		private void initExcelTemplate()
		{

			sheet = workbook.GetSheetAt(sheetIndex);
			readExcelTemplateDatasIndenfier();
			readExcelTemplateStyles();

			readExcelTemplateOrder();

			//获取最后一行位置
			lastRowIndex = sheet.LastRowNum;

			if (!isOrder) //无序号列
						  //创建初始化后的第一个当前数据行
				currentRow = sheet.CreateRow(currentDataCellRowIndex);


		}

		/// <summary>
		/// 初始化多工作表
		/// </summary>
		/// <param name="sheetIndex"></param>
		public void initExcelTemplate(int sheetIndex)
		{

			sheet = workbook.GetSheetAt(sheetIndex);
			readExcelTemplateDatasIndenfier();
			readExcelTemplateStyles();

			readExcelTemplateOrder();

			//获取最后一行位置
			lastRowIndex = sheet.LastRowNum;

			if (!isOrder) //无序号列
						  //创建初始化后的第一个当前数据行
				currentRow = sheet.CreateRow(currentDataCellRowIndex);


		}

		/// <summary>
		/// 读取Excel模板中的Datas的标识符并初始化
		/// </summary>
		/// <exception cref="NpoiException"></exception>
		private void readExcelTemplateDatasIndenfier()
		{

			bool ifFindDatasIndentifier = false;

			foreach (IRow row in sheet)
			{

				if (ifFindDatasIndentifier)
				{

					break;
				}

				foreach (ICell cell in row)
				{

					//判断datas标识符所在位置
					if (cell.CellType != CellType.String) continue;
					//记录datas的标识符
					if (cell.StringCellValue.Trim().Contains(DATACELL_INDENTIFIER))
					{

						startDataCellColIndex = cell.ColumnIndex;
						startDataCellRowIndex = cell.RowIndex;
						currentDataCellColIndex = startDataCellColIndex;
						currentDataCellRowIndex = startDataCellRowIndex;


						ifFindDatasIndentifier = true;

						break;
					}

				}

			}

			if (!ifFindDatasIndentifier)
			{

				throw new NpoiException("模板中没有datas标识存在");
			}


			//Console.WriteLine(DATACELL_INDENTIFIER+"位置:("+startDataCellRowIndex+","+startDataCellColIndex+")");
		}


		/// <summary>
		/// 读取ExcelTemplate模板的样式以及记录样式
		/// </summary>
		public void readExcelTemplateStyles()
		{

			bool ifFindDefaultStyles = false;

			foreach (IRow row in sheet)
			{

				foreach (ICell cell in row)
				{

					if (cell.CellType != CellType.String) continue;

					/*
					 * 不在使用，跟随单元格的默认样式
					 * if(cell.getStringCellValue().trim().contains(DATACELL_INDENTIFIER)){
						//获取数据开始的单元格的样式
						defaultCellStyle=cell.getCellStyle();
						//默认行高
						defaultRowHeight=row.getHeightInPoints();
					}*/

					//记录默认样式
					if (cell.StringCellValue.Trim().Contains(DEFAULT_STYLE_INDENTIFIER) && !ifFindDefaultStyles)
					{

						//获取数据开始的单元格的样式
						defaultCellStyle = cell.CellStyle;
						//默认行高
						defaultRowHeight = row.HeightInPoints;

						ifFindDefaultStyles = true;

					}
					else
					{

						if (!ifFindDefaultStyles)
						{
							//没有找到就设置默认样式以单元格设置样式了
							//获取数据开始的单元格的样式
							defaultCellStyle = cell.CellStyle;
							//默认行高
							defaultRowHeight = row.HeightInPoints;

						}

					}

					if (cell.StringCellValue.Trim().Contains(STYLES_INDENTIFIER))
					{

						styleMaps.Add(cell.ColumnIndex, cell.CellStyle);
					}

				}

			}

		}

		/// <summary>
		/// 替换一些Excel模板中的一些常量例如以#开头的title/date/department等，可以增加
		/// </summary>
		/// <param name="datas"></param>
		public void replaceExcelTemplate(Dictionary<string, object> datas)
		{

			if (datas == null) return;

			foreach (IRow row in sheet)
			{

				foreach (ICell cell in row)
				{

					if (cell.CellType != CellType.String) continue;

					//查找以#开头的或包含(如果要替换的数据出现在中间，可能需要修改规则例如用#...#约束，可考虑正则判断)
					if (cell.StringCellValue.Trim().StartsWith("#") || cell.StringCellValue.Trim().Contains("#"))
					{

						//获得模板上的名称
						string tempName = cell.StringCellValue;

						//遍历键集,重组
						Dictionary<string, object> newInfos = new Dictionary<string, object>();
						foreach (string key in datas.Keys)
						{
							//获取值
							object keyValue = "";
							datas.TryGetValue(key, out keyValue);
							//重组键
							string newKey = "#" + key;
							newInfos.Add(newKey, keyValue);
						}

						//统计#个数
						int c1 = 0;
						for (int i = 0; i < tempName.Length; i++)
						{
							if (tempName[i] == '#')
							{
								c1++;
							}
						}

						if (c1 == 1)
						{
							if (datas.ContainsKey(tempName.Substring(tempName.IndexOf("#") + 1).Trim()))
							{

								//Console.WriteLine(tempName.Substring(tempName.IndexOf("#")+1));
								object dataValue = "";
								datas.TryGetValue(tempName.Substring(tempName.IndexOf("#") + 1), out dataValue);

								//根据类型把obj转为相应的数据
								Type dvType = dataValue.GetType();
								if (dvType == typeof(int))
								{
									cell.SetCellValue(int.Parse(dataValue.ToString()));
								}
								else if (dvType == typeof(double))
								{
									cell.SetCellValue(double.Parse(dataValue.ToString()));
								}
								else
								{

									cell.SetCellValue(cell.StringCellValue.Replace(tempName.Substring(tempName.IndexOf("#")), dataValue.ToString()));
								}

							}
						}
						else
						{
							//遍历重组infos
							foreach (string key in newInfos.Keys)
							{

								if (tempName.Contains(key))
								{ //用value替换掉包含的key的内容

									object value = "";
									newInfos.TryGetValue(key, out value);
									tempName = tempName.Replace(key, value.ToString());
								}

							}
							//重设该Cell内容
							cell.SetCellValue(tempName);
						}

					}
				}
			}

		}


		/// <summary>
		/// 读取Excel模板的order标识符的位置
		/// </summary>
		public void readExcelTemplateOrder()
		{

			foreach (IRow row in sheet)
			{

				if (ifFindOrder) break;

				foreach (ICell cell in row)
				{

					if (cell.CellType != CellType.String) continue;

					if (cell.StringCellValue.Trim().Contains(ORDER_INDENTIFIER))
					{

						insertOrderColIndex = cell.ColumnIndex;
						insertOrderRowIndex = cell.RowIndex;


						ifFindOrder = true;

						break;

					}
				}
			}

			if (!ifFindOrder)
			{
				//如果没找到序号标识模板中，就设置为无序号

				//			isOrder=false;

				if (isOrder)
				{
					//模板中无序号标识，用户还是需要设置序号

					//设置此时序号应该在的位置 应该在datas标识符的位置
					insertOrderRowIndex = startDataCellRowIndex;
					insertOrderColIndex = startDataCellColIndex;

					//数据单元格整体向后移动
					currentDataCellColIndex = startDataCellColIndex + 1;
				}

				//Console.WriteLine("excel模板中没有order序号标识,如果设置序号可能会出现错误");
			}



			if (!isOrder && ifFindOrder)
			{//不显示序号 设置为不显示序号且模板中有order的表示才整体移动

				//数据单元格列整体向前移动
				currentDataCellColIndex = startDataCellColIndex - 1;
			}

		}

		/// <summary>
		/// 插入序号默认序号从1开始
		/// </summary>
		public void insertOrder()
		{

			insertOrder(1, insertOrderRowIndex);
		}



		/// <summary>
		/// 插入序号
		/// </summary>
		/// <param name="startIndex">开始序号</param>
		/// <param name="insertOrderStartRowIndex">插入序号的开始行下标</param>
		public void insertOrder(int startIndex, int insertOrderStartRowIndex)
		{

			if (isOrder && insertOrderColIndex != -1 && insertOrderStartRowIndex != -1)
			{//设置序号


				for (int i = insertOrderStartRowIndex; i <= currentDataCellRowIndex; i++)
				{

					//获取行 创建单元格
					ICell orderCell = sheet.GetRow(i).CreateCell(insertOrderColIndex);
					if (orderCell != null)
					{

						//应用样式
						setStyles(orderCell);
						orderCell.SetCellValue(startIndex);

					}
					startIndex++;
				}


			}
			else
			{//删除序号列



			}
		}

		/// <summary>
		/// 填充表格
		/// </summary>
		/// <param name="cellValue">表格值</param>
		/// <param name="type">类型</param>
		private void inputRowCell(object cellValue, TYPE type)
		{


			//自动创建下一行
			if (currentDataCellColIndex == excelColNums - 1 && !isOrder)
			{//到达一行的最后一列

				createNextRow();
			}
			else if (currentDataCellColIndex == excelColNums && isOrder)
			{//到达一行的最后一列

				createNextRow();
			}

			//根据当前行根据当前列创建单元格
			ICell newCell = currentRow.CreateCell(currentDataCellColIndex);

			if (newCell == null) return;
			setStyles(newCell);

			//数据为空
			if (cellValue == null || cellValue.ToString().Equals(""))
			{

				newCell.SetCellValue(nullPlaceHolder);
			}

			switch (type)
			{
				case TYPE.STring:
					//设置值
					newCell.SetCellValue(cellValue.ToString());


					break;
				case TYPE.BOolean:

					newCell.SetCellValue(bool.Parse(cellValue.ToString()));
					break;
				case TYPE.DAte:

					newCell.SetCellValue(DateTime.Parse(cellValue.ToString()));

					break;
				case TYPE.DOuble:

					newCell.SetCellValue(Double.Parse(cellValue.ToString()));

					break;
				case TYPE.INteger:

					newCell.SetCellValue(int.Parse(cellValue.ToString()));

					break;
				//Calendar目前没有合适处理
				/*case TYPE.CAlendar:
					newCell.SetCellValue((Calendar)cellValue.ToString);
					break;*/

				default:
					break;
			}




			currentDataCellColIndex++;
		}

		/// <summary>
		/// 数据输入到一行的所有单元格
		/// </summary>
		/// <param name="cellValue"></param>
		public void inputRowCell(string cellValue)
		{


			inputRowCell(cellValue, TYPE.STring);
		}
		
		public void inputRowCell(int cellValue)
		{

			inputRowCell(cellValue, TYPE.INteger);

		}


		public void inputRowCell(bool cellValue)
		{

			inputRowCell(cellValue, TYPE.BOolean);

		}


		public void inputRowCell(DateTime cellValue, string pattern)
		{


			SimpleDateFormat sdf = new SimpleDateFormat(pattern);

			inputRowCell(sdf.Format(cellValue));

		}

		/// <summary>
		/// 默认转换模式为 yyyy-MM-dd
		/// </summary>
		/// <param name="cellValue"></param>
		public void inputRowCell(DateTime cellValue)
		{

			inputRowCell(cellValue, excelDatePattern.getDatePattern());

		}

		
		public void inputRowCell(double cellValue)
		{

			inputRowCell(cellValue, TYPE.DOuble);

		}

		/*///<summary></summary>
		/// 当前日期使用 目前没有合适处理
		///<summary></summary>
		public void inputRowCell(Calendar cellValue, ExcelDatePattern ec)
		{
			SimpleDateFormat sdf = new SimpleDateFormat();
			DateTime date = cellValue.ToDateTime(cellValue.GetYear,cellValue.GetMonth,cellValue.GetDaysInYear,cellValue.GetHour,cellValue.GetMinute,cellValue.GetSecond,cellValue.GetMilliseconds);
			inputRowCell(sdf.Format(date));
		}*/

		/*///<summary></summary>
		/// 当前日期使用 目前没有合适处理
		///<summary></summary>
		public void inputRowCell(Calendar cellValue)
		{
			inputRowCell(cellValue, ExcelDatePattern.getInstance());
		}*/

		///<summary></summary>
		///为单元格设置样式
		///<summary></summary>
		private void setStyles(ICell newCell)
		{



			//不包含styles的列,设置为默认样式
			if (styleMaps.ContainsKey(newCell.ColumnIndex))
			{

				ICellStyle cellStyle = null;
				styleMaps.TryGetValue(newCell.ColumnIndex, out cellStyle);

				if (cellStyle != null)
					newCell.CellStyle = cellStyle;

			}
			else
			{

				if (defaultCellStyle != null)
				{

					//设置默认样式
					newCell.CellStyle = defaultCellStyle;
					newCell.Row.HeightInPoints = defaultRowHeight;

				}
			}
		}

		/// <summary>
		/// 通过Excel模板写出数据到一个Excel文件中
		/// </summary>
		/// <param name="path"></param>
		/// <exception cref="NpoiException"></exception>
		/// <exception cref="IOException"></exception>
		public void writeDataToExcel(string path)
		{


			FileStream fs = null;
			BufferedStream bs = null;

			//判断文件夹是否存在 不存在则创建
			string tempPath = System.IO.Path.GetDirectoryName(path);
			if (!System.IO.Directory.Exists(tempPath))
				System.IO.Directory.CreateDirectory(tempPath);
			try
			{
				fs = new FileStream(path, FileMode.Create);
				bs = new BufferedStream(fs);

				workbook.Write(bs);
			}
			catch (FileNotFoundException e)
			{
				
				throw new NpoiException("写入的文件不存在或正在使用");
			}
			catch (IOException e)
			{
				
				throw new NpoiException("写入数据失败" + e.Message);

			}
			finally
			{

				try
				{
					if (bs != null) bs.Close();

					if (fs != null)
						fs.Close();
				}
				catch (IOException e)
				{
			
					throw new IOException(e.Message);
				}

			}
		}

		/// <summary>
		/// 通过Excel模板写出数据到输出流
		/// </summary>
		/// <param name="stream"></param>
		/// <exception cref="NpoiException"></exception>
		public void writeDataToStream(Stream stream)
		{

			try
			{
				workbook.Write(stream);
			}
			catch (IOException e)
			{
				

				throw new NpoiException("写入流失败" + e.ToString());
			}
			finally
			{

				try
				{

					if (stream != null) stream.Close();

				}
				catch (Exception e)
				{

				}
			}
		}


		/// <summary>
		/// 一行的所有单元格填满之后创建下一行
		/// </summary>
		public void createNextRow()
		{

			//移动行
			moveRow(currentDataCellRowIndex + 1, lastRowIndex, 1);

			//创建一行
			currentDataCellRowIndex++;
			currentRow = sheet.CreateRow(currentDataCellRowIndex);

			//列初始化
			if (!isOrder && ifFindOrder) //存在order并且设置为不显示序号单元格列才整体移动
				currentDataCellColIndex = startDataCellColIndex - 1;
			else currentDataCellColIndex = startDataCellColIndex;

			if (!ifFindOrder && isOrder)
			{//模板中找不到order标识，用户还是想要设置

				//数据列整体后移一位
				currentDataCellColIndex = startDataCellColIndex + 1;
			}
		}


		/// <summary>
		/// 移动行
		/// </summary>
		/// <param name="startRow"></param>
		/// <param name="endRow"></param>
		/// <param name="n">移动n行</param>
		public void moveRow(int startRow, int endRow, int n)
		{

			//最后一行大于当前行
			if (lastRowIndex > currentDataCellRowIndex)
			{

				sheet.ShiftRows(startRow, endRow, n, true, true);
				//最后一行++
				lastRowIndex++;
			}

		}

		public void setOrderIndentifier()
		{
			ICell orderNameCell = null;
			if (isOrder)
			{//有序号列获取当前行

				if (ifFindOrder)//查找到order标识
				{
					currentRow = sheet.GetRow(insertOrderRowIndex);
					orderNameCell = currentRow.GetCell(insertOrderColIndex);
				}
				else 
				{ //创建序号行
				
					currentRow = sheet.CreateRow(insertOrderRowIndex);
					orderNameCell = currentRow.CreateCell(insertOrderColIndex);
				}

				orderNameCell.SetCellValue(orderName);
				//设置样式
				setStyles(orderNameCell);

				//如果开始数据的行 datas存在 当前行设置为数据所在行
				if(startDataCellRowIndex!=-1 && startDataCellRowIndex!=currentRow.RowNum)
					currentRow = sheet.CreateRow(startDataCellRowIndex);
			}

		}



	}

}
