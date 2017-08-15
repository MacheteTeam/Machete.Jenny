using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Machete.Jenny
{
    public static class ExcelHelper
    {
        /// <summary>
        /// 导入Excel为  TList  支持自定义列
        /// </summary>
        /// <typeparam name="T">数据类型</typeparam>
        /// <param name="strFileName"><文件名/param>
        /// <param name="sheetIndex">需要读取的Sheet页</param>
        /// <param name="titleDictionaries">自定义数据列</param>
        /// <returns></returns>
        public static List<T> ExcelToList<T>(Type instanceType, string strFileName, int sheetIndex = 0, bool haveTitles = false, Dictionary<string, string> titleDic = null) where T : class, new()
        {
            PropertyInfo[] myPropertyInfo = instanceType.GetProperties(BindingFlags.Public | BindingFlags.Instance);  //获取所有属性
            List<T> tList = new List<T>();
            var propertys = typeof(T).GetProperties();
            HSSFWorkbook hssfworkbook = null;
            XSSFWorkbook xssfworkbook = null;
            string fileExt = Path.GetExtension(strFileName);    //获取文件的后缀名
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                if (fileExt == ".xls")
                    hssfworkbook = new HSSFWorkbook(file);
                else if (fileExt == ".xlsx")
                    xssfworkbook = new XSSFWorkbook(file);     //初始化太慢了，不知道这是什么bug
            }
            if (hssfworkbook != null)
            {
                HSSFSheet sheet = (HSSFSheet)hssfworkbook.GetSheetAt(sheetIndex);
                if (sheet != null)
                {
                    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                    if (haveTitles)
                    {
                        HSSFRow headerRow = (HSSFRow)sheet.GetRow(0);
                        int cellCount = headerRow.LastCellNum;
                        for (int j = 0; j < cellCount; j++)
                        {
                            HSSFCell cell = (HSSFCell)headerRow.GetCell(j);   //获取Excel标题
                        }
                    }
                    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                    {
                        HSSFRow row = (HSSFRow)sheet.GetRow(i);

                        var obj = new T();
                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                        {
                            string cellValue = "";
                            if (row.GetCell(j) != null)
                            {
                                cellValue = row.GetCell(j).ToString();
                                string dkey = titleDic.Keys.ToArray()[j];
                                string dataType = (myPropertyInfo[j].PropertyType).FullName;  //获取数据类型
                                foreach (var p in propertys)
                                {
                                    string name = p.Name;
                                    if (name == dkey)
                                    {
                                        if (dataType == "System.String")
                                        {
                                            p.SetValue(obj, cellValue, null);
                                        }
                                        else if (dataType == "System.DateTime")
                                        {
                                            DateTime pdt = Convert.ToDateTime(cellValue);
                                            p.SetValue(obj, pdt, null);
                                        }
                                        else if (dataType == "System.Boolean")
                                        {
                                            bool pb = Convert.ToBoolean(cellValue);
                                            p.SetValue(obj, pb, null);
                                        }
                                        else if (dataType == "System.Int16")
                                        {
                                            Int16 pi16 = Convert.ToInt16(cellValue);
                                            p.SetValue(obj, pi16, null);
                                        }
                                        else if (dataType == "System.Int32")
                                        {
                                            Int32 pi32 = Convert.ToInt32(cellValue);
                                            p.SetValue(obj, pi32, null);
                                        }
                                        else if (dataType == "System.Int64")
                                        {
                                            Int64 pi64 = Convert.ToInt64(cellValue);
                                            p.SetValue(obj, pi64, null);
                                        }
                                        else if (dataType == "System.Byte")
                                        {
                                            Byte pb = Convert.ToByte(cellValue);
                                            p.SetValue(obj, pb, null);
                                        }
                                        else if (dataType == "System.Decimal")
                                        {
                                            System.Decimal pd = Convert.ToDecimal(cellValue);
                                            p.SetValue(obj, pd, null);
                                        }
                                        else if (dataType == "System.Double")
                                        {
                                            double pd = Convert.ToDouble(cellValue);
                                            p.SetValue(obj, pd, null);
                                        }
                                        else
                                        {
                                            p.SetValue(obj, null, null);
                                        }
                                    }
                                }
                            }
                        }
                        tList.Add(obj);
                    }
                }
            }
            else if (xssfworkbook != null)
            {
                XSSFSheet xSheet = (XSSFSheet)xssfworkbook.GetSheetAt(sheetIndex);
                if (xSheet != null)
                {
                    System.Collections.IEnumerator rows = xSheet.GetRowEnumerator();
                    if (haveTitles)
                    {
                        XSSFRow headerRow = (XSSFRow)xSheet.GetRow(0);
                        int cellCount = headerRow.LastCellNum;
                        for (int j = 0; j < cellCount; j++)
                        {
                            XSSFCell cell = (XSSFCell)headerRow.GetCell(j);   //获取Excel标题
                        }
                    }
                    for (int i = (xSheet.FirstRowNum + 1); i <= xSheet.LastRowNum; i++)
                    {
                        XSSFRow row = (XSSFRow)xSheet.GetRow(i);

                        var obj = new T();
                        for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
                        {
                            string value = "";
                            if (row.GetCell(j) != null)
                            {
                                value = row.GetCell(j).ToString();
                                string dkey = titleDic.Keys.ToArray()[j];
                                string dvalue = titleDic.Values.ToArray()[j];
                                string str = (myPropertyInfo[j].PropertyType).FullName;  //获取数据类型
                                foreach (var p in propertys)
                                {
                                    string name = p.Name;
                                    if (name == dkey)
                                    {
                                        if (str == "System.String")
                                        {
                                            p.SetValue(obj, value, null);
                                        }
                                        else if (str == "System.DateTime")
                                        {
                                            DateTime pdt = Convert.ToDateTime(value);
                                            p.SetValue(obj, pdt, null);
                                        }
                                        else if (str == "System.Boolean")
                                        {
                                            bool pb = Convert.ToBoolean(value);
                                            p.SetValue(obj, pb, null);
                                        }
                                        else if (str == "System.Int16")
                                        {
                                            Int16 pi16 = Convert.ToInt16(value);
                                            p.SetValue(obj, pi16, null);
                                        }
                                        else if (str == "System.Int32")
                                        {
                                            Int32 pi32 = Convert.ToInt32(value);
                                            p.SetValue(obj, pi32, null);
                                        }
                                        else if (str == "System.Int64")
                                        {
                                            Int64 pi64 = Convert.ToInt64(value);
                                            p.SetValue(obj, pi64, null);
                                        }
                                        else if (str == "System.Byte")
                                        {
                                            Byte pb = Convert.ToByte(value);
                                            p.SetValue(obj, pb, null);
                                        }
                                        else if (str == "System.Decimal")
                                        {
                                            System.Decimal pd = Convert.ToDecimal(value);
                                            p.SetValue(obj, pd, null);
                                        }
                                        else if (str == "System.Double")
                                        {
                                            double pd = Convert.ToDouble(value);
                                            p.SetValue(obj, pd, null);
                                        }
                                        else
                                        {
                                            p.SetValue(obj, null, null);
                                        }
                                    }
                                }
                            }
                        }
                        tList.Add(obj);
                    }
                }
            }
            return tList;
        }



        #region ListToExcel<T>(List<T> myList, string saveFileName = null, bool isOpen = false,string saveFilePath = null, string strHeaderText = null, Dictionary<string, string> titleDictionaries = null) List导出到Excel文件--C/S

        /// <summary>
        /// C/S List导出数据到Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="myList">需要导出的泛型List</param>
        /// <param name="saveFileName">保存的文件名称，默认没有，调用的时候最好加上，中英文都支持</param>
        /// <param name="isOpen">导出后是否打开文件和所在文件夹</param>
        /// <param name="saveFilePath">默认保存在“我的文档”中，可自定义保存的文件夹路径</param>
        /// <param name="strHeaderText">Excel中第一行的标题文字，默认没有，可以自定义</param>
        /// <param name="titleDictionaries">Excel中需要导出的列的字典映射，默认绑定List的列名</param>
        public static void ListToExcel<T>(List<T> myList, string saveFileName = null, bool isOpen = false,
                    string saveFilePath = null, string strHeaderText = null, Dictionary<string, string> titleDic = null)
        {
            using (MemoryStream ms = ListToExcel(myList, strHeaderText, titleDic))
            {
                if (string.IsNullOrEmpty(saveFileName)) //文件名为空
                {
                    saveFileName = DateTime.Now.Ticks.ToString();
                }
                if (string.IsNullOrEmpty(saveFilePath) || !Directory.Exists(saveFilePath)) //保存路径为空或者不存在
                {
                    saveFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments); //默认在文档文件夹中
                }
                string saveFullPath = saveFilePath + "\\" + saveFileName + ".xls";
                if (File.Exists(saveFullPath)) //验证文件重复性
                {
                    saveFullPath = saveFilePath + "\\" + saveFileName +
                                   DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss").Replace(":", "-").Replace(" ", "-") +
                                   ".xls";
                }
                using (FileStream fileStream = new FileStream(saveFullPath, FileMode.Create, FileAccess.Write))
                {
                    byte[] data = ms.ToArray();
                    fileStream.Write(data, 0, data.Length);
                    fileStream.Flush();
                }
                if (isOpen)
                {
                    Process.Start(saveFilePath); //打开文件夹
                    Process.Start(saveFullPath); //打开文件
                }
            }
        }

        #endregion

        /// <summary>
        /// List导出到Excel的MemoryStream
        /// </summary>
        /// <param name="list">需要导出的泛型List</param>
        /// <param name="strHeaderText">第一行标题头</param>
        /// <param name="titleDictionaries">列名称字典映射</param>
        /// <returns>内存流</returns>
        private static MemoryStream ListToExcel<T>(List<T> list, string strHeaderText = null,
            Dictionary<string, string> titleDic = null)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet();

            #region 右击文件 属性信息

            {
                DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                dsi.Company = "NPOI";
                workbook.DocumentSummaryInformation = dsi;

                SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                si.Author = "文件作者信息"; //填加xls文件作者信息
                si.ApplicationName = "创建程序信息"; //填加xls文件创建程序信息
                si.LastAuthor = "最后保存者信息"; //填加xls文件最后保存者信息
                si.Comments = "作者信息"; //填加xls文件作者信息
                si.Title = "标题信息"; //填加xls文件标题信息
                si.Subject = "主题信息"; //填加文件主题信息
                si.CreateDateTime = DateTime.Now;
                workbook.SummaryInformation = si;
            }

            #endregion

            HSSFCellStyle dateStyle = (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFDataFormat format = (HSSFDataFormat)workbook.CreateDataFormat();
            dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

            //取得列宽
            //通过反射得到对象的属性集合  
            PropertyInfo[] myPropertyInfo = list.First().GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
            int fieldsCount = myPropertyInfo.Length;
            int[] arrColWidth = new int[fieldsCount];
            int index = 0;
            foreach (var item in myPropertyInfo)
            {
                arrColWidth[index] = Encoding.GetEncoding(936).GetBytes(item.Name).Length;
                index++;
            }
            for (int i = 0; i < list.Count; i++)
            {
                for (int j = 0; j < fieldsCount; j++)
                {
                    int intTemp = Encoding.GetEncoding(936).GetBytes((myPropertyInfo[j].GetValue(list[i])).ToString()).Length;
                    if (intTemp > arrColWidth[j])
                    {
                        arrColWidth[j] = intTemp;
                    }
                }
            }
            int rowIndex = 0;
            bool flag = false;
            for (int row = -1; row < list.Count; row++)
            {
                #region 新建表，填充表头，填充列头，样式

                if (rowIndex == 65535 || rowIndex == 0)
                {
                    if (rowIndex != 0)
                    {
                        sheet = (HSSFSheet)workbook.CreateSheet();
                    }

                    #region 表头及样式

                    if (!string.IsNullOrEmpty(strHeaderText))
                    {
                        HSSFRow headerRow = (HSSFRow)sheet.CreateRow(0);
                        headerRow.HeightInPoints = 25;
                        headerRow.CreateCell(0).SetCellValue(strHeaderText);

                        HSSFCellStyle headStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                        HSSFFont font = (HSSFFont)workbook.CreateFont();
                        font.FontHeightInPoints = 20;
                        font.Boldweight = 700;
                        font.FontName = "宋体";
                        headStyle.SetFont(font);
                        headerRow.GetCell(0).CellStyle = headStyle;
                        rowIndex++;
                    }

                    #endregion


                    #region 列头及样式
                    {
                        HSSFRow headerRow = (HSSFRow)sheet.CreateRow(rowIndex);
                        HSSFCellStyle headStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                        HSSFFont font = (HSSFFont)workbook.CreateFont();
                        font.FontHeightInPoints = 14;
                        font.Boldweight = 500;
                        font.FontName = "宋体";
                        headStyle.SetFont(font);

                        Type myType = list[0].GetType();
                        int cn = 0;
                        if (titleDic.Count != 0)
                        {
                            foreach (string titleDictionary in titleDic.Keys)
                            {
                                PropertyInfo p = myType.GetProperty(titleDictionary);
                                if (p != null)
                                {
                                    headerRow.CreateCell(cn).SetCellValue(titleDic[titleDictionary]);
                                    headerRow.GetCell(cn).CellStyle = headStyle;
                                    //设置列宽
                                    sheet.SetColumnWidth(cn, (arrColWidth[cn] + 1) * 500);
                                }
                                cn++;
                            }
                        }
                        else
                        {
                            int colIndex = 0;
                            foreach (var column in myPropertyInfo)
                            {
                                headerRow.CreateCell(colIndex).SetCellValue(column.Name);
                                headerRow.GetCell(colIndex).CellStyle = headStyle;
                                //设置列宽
                                sheet.SetColumnWidth(colIndex, (arrColWidth[colIndex] + 1) * 500);
                                colIndex++;
                            }
                        }
                        rowIndex++;
                    }

                    #endregion
                }

                #endregion

                #region 填充内容

                HSSFRow dataRow = (HSSFRow)sheet.CreateRow(rowIndex);
                if (flag)
                {
                    for (int j = 0; j < fieldsCount; j++)
                    {
                        HSSFCell newCell = (HSSFCell)dataRow.CreateCell(j);

                        if (titleDic.ContainsKey((myPropertyInfo[j]).Name))
                        {
                            string drValue = (myPropertyInfo[j].GetValue(list[row])).ToString();
                            switch (((myPropertyInfo[j]).PropertyType).FullName)
                            {
                                case "System.String": //字符串类型
                                    newCell.SetCellValue(drValue);
                                    break;
                                case "System.DateTime": //日期类型
                                    DateTime dateV;
                                    DateTime.TryParse(drValue, out dateV);
                                    newCell.SetCellValue(dateV);

                                    newCell.CellStyle = dateStyle; //格式化显示
                                    break;
                                case "System.Boolean": //布尔型
                                    bool boolV = false;
                                    bool.TryParse(drValue, out boolV);
                                    newCell.SetCellValue(boolV);
                                    break;
                                case "System.Int16": //整型
                                case "System.Int32":
                                case "System.Int64":
                                case "System.Byte":
                                    int intV = 0;
                                    int.TryParse(drValue, out intV);
                                    newCell.SetCellValue(intV);
                                    break;
                                case "System.Decimal": //浮点型
                                case "System.Double":
                                    double doubV = 0;
                                    double.TryParse(drValue, out doubV);
                                    newCell.SetCellValue(doubV);
                                    break;
                                case "System.DBNull": //空值处理
                                    newCell.SetCellValue("");
                                    break;
                                default:
                                    newCell.SetCellValue("");
                                    break;
                            }
                        }
                    }
                }
                else
                {
                    rowIndex--;
                }
                flag = true;

                #endregion

                rowIndex++;
            }
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
                //sheet.Dispose();
                //workbook.Dispose();//一般只用写这一个就OK了，他会遍历并释放所有资源，但当前版本有问题所以只释放sheet
                return ms;
            }
        }
    }
}
