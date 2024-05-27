using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using System;
using System.Data;
using System.IO;

namespace Fountain.WinConsole.NPOIDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                DataTable productTable = CreateProductInfo();
                // 创建工作簿对象
                IWorkbook workbook = new SXSSFWorkbook();
                // 创建工作簿的工作表对象
                ISheet sheet = workbook.CreateSheet("Product");
                // 创建单元格样式
                ICellStyle cellTitleStyle = workbook.CreateCellStyle();
                // 创建字体对象
                IFont font = workbook.CreateFont();
                font.FontName = "微软雅黑";
                // 字体颜色
                font.Color = IndexedColors.Black.Index;
                // 是否斜体
                font.IsItalic = false;
                // 是否粗体
                font.IsBold = true;
                // 下划线
                font.Underline = FontUnderlineType.None;
                // 字体大小
                font.FontHeightInPoints = 12;
                // 样式绑定到单元格
                cellTitleStyle.SetFont(font);
                // 创建标题行
                IRow row = sheet.CreateRow(0);
                // 根据数据标字段给标题行设置值
                for (int i = 0; i < productTable.Columns.Count; i++)
                {
                    // 创建行对象单元格
                    ICell cell = row.CreateCell(i);
                    // 设置单元格内容
                    cell.SetCellValue(productTable.Columns[i].ColumnName);
                    // 绑定样式
                    cell.CellStyle = cellTitleStyle; 
                }
                // 创建单元格样式
                ICellStyle cellStyleContent = workbook.CreateCellStyle();
                // 创建字体对象
                IFont fontContent = workbook.CreateFont();
                fontContent.FontName = "微软雅黑";
                // 字体颜色
                fontContent.Color = IndexedColors.Black.Index;
                // 是否斜体
                fontContent.IsItalic = false;
                // 是否粗体
                fontContent.IsBold = false;
                // 下划线
                fontContent.Underline = FontUnderlineType.None;
                // 字体大小
                fontContent.FontHeightInPoints = 9;
                // 样式绑定到单元格
                cellStyleContent.SetFont(fontContent);

                // 创建单元格样式
                IDataFormat dataFormat = workbook.CreateDataFormat();

                // 将表数据导出
                int rowIndex = 1;
                foreach (DataRow rowItem in productTable.Rows)
                {
                    IRow rowFill = sheet.CreateRow(rowIndex);
                    for (int j = 0; j < productTable.Columns.Count; j++)
                    {
                        // 创建行对象单元格
                        ICell cell = rowFill.CreateCell(j);
                       
                        Type dataType = productTable.Columns[j].DataType;
                        switch (dataType.Name)
                        {
                            case "Int32":
                                cellStyleContent.DataFormat = dataFormat.GetFormat("###");
                                cellStyleContent.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
                                // 设置单元格内容
                                cell.SetCellValue(Convert.ToInt32(rowItem[productTable.Columns[j].ColumnName]));
                                break;
                            case "Double":
                                cellStyleContent.DataFormat = dataFormat.GetFormat("###.00");
                                cellStyleContent.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
                                // 设置单元格内容
                                cell.SetCellValue(Convert.ToDouble(rowItem[productTable.Columns[j].ColumnName]));
                                break;
                            default:
                                cellStyleContent.DataFormat = dataFormat.GetFormat("General");
                                // 设置单元格内容
                                cell.SetCellValue(rowItem[productTable.Columns[j].ColumnName].ToString());
                                break;
                        }
                        // 绑定样式
                        cell.CellStyle = cellStyleContent;
                    }
                    rowIndex++;
                }
                // 自适应列宽
                
                // 定义储存文件名与位置
                string excelFileName = string.Format("{0}{1}", AppDomain.CurrentDomain.BaseDirectory, "ProductInfo.xlsx");
                // 将工作簿对象写入文件
                using (FileStream fileStream = new FileStream(excelFileName, FileMode.Create, FileAccess.ReadWrite))
                {
                    workbook.Write(fileStream);
                }
                // 关闭工作簿对象
                workbook.Close();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
            // 等待控制台输入
            Console.ReadKey();
        }
        /// <summary>
        /// 生成表数据
        /// </summary>
        /// <returns></returns>
        private static DataTable CreateProductInfo()
        {
            DataTable productTabel = null;
            try
            {
                productTabel = new DataTable();
                productTabel.TableName = "Product";
                // 构建表结构
                productTabel.Columns.Add("ProductID", typeof(string));
                productTabel.Columns.Add("Barcode", typeof(string));
                productTabel.Columns.Add("ProductName", typeof(string));
                productTabel.Columns.Add("CategoryID", typeof(string));
                productTabel.Columns.Add("Price", typeof(double));
                // 生成表数据
                for (int i = 0; i < 5; i++)
                {
                    DataRow rowItem = productTabel.NewRow();
                    rowItem["ProductID"] = "A499797880-" + i.ToString().PadLeft(2, '0');
                    rowItem["Barcode"] = "A499797880" + i.ToString().PadLeft(2, '0');
                    rowItem["ProductName"] = "裤子";
                    rowItem["CategoryID"] = "P";
                    rowItem["Price"] = 139.5 * (i + 0.56);
                    productTabel.Rows.Add(rowItem);
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
            }
            return productTabel;
        }
    }
}
