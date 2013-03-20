using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using Microsoft.Office;
using Microsoft.Office.Core;
namespace OADAL
{
    /// <summary>
    /// ExcelOperate 的摘要说明。Excel操作函数
    /// </summary>
    public class ExcelOperate
    {
        private object mValue = System.Reflection.Missing.Value;

        public ExcelOperate()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        public void Merge(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell)
        {
            CurSheet.get_Range(objStartCell, objEndCell).Merge(mValue);
        }
        /// <summary>
        /// 设置连续区域的字体大小
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="strStartCell">开始单元格</param>
        /// <param name="strEndCell">结束单元格</param>
        /// <param name="intFontSize">字体大小</param>
        public void SetFontSize(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell, int intFontSize)
        {
            CurSheet.get_Range(objStartCell, objEndCell).Font.Size = intFontSize.ToString();
        }

        /// <summary>
        /// 横向打印
        /// </summary>
        /// <param name="CurSheet"></param>
        public void xlLandscape(Microsoft.Office.Interop.Excel._Worksheet CurSheet)
        {
            CurSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;

        }
        /// <summary>
        /// 纵向打印
        /// </summary>
        /// <param name="CurSheet"></param>
        public void xlPortrait(Microsoft.Office.Interop.Excel._Worksheet CurSheet)
        {
            CurSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
        }


        /// <summary>
        /// 在指定单元格插入指定的值
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="Cell">单元格 如Cells[1,1]</param>
        /// <param name="objValue">文本、数字等值</param>
        public void WriteCell(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objCell, object objValue)
        {
            CurSheet.get_Range(objCell, mValue).Value2 = objValue;

        }

        /// <summary>
        /// 在指定Range中插入指定的值
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="StartCell">开始单元格</param>
        /// <param name="EndCell">结束单元格</param>
        /// <param name="objValue">文本、数字等值</param>
        public void WriteRange(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell, object objValue)
        {
            CurSheet.get_Range(objStartCell, objEndCell).Value2 = objValue;
        }

        /// <summary>
        /// 合并单元格，并在合并后的单元格中插入指定的值
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        /// <param name="objValue">文本、数字等值</param>
        public void WriteAfterMerge(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell, object objValue)
        {
            CurSheet.get_Range(objStartCell, objEndCell).Merge(mValue);
            CurSheet.get_Range(objStartCell, mValue).Value2 = objValue;

        }

        /// <summary>
        /// 为单元格设置公式
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objCell">单元格</param>
        /// <param name="strFormula">公式</param>
        public void SetFormula(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objCell, string strFormula)
        {
            CurSheet.get_Range(objCell, mValue).Formula = strFormula;
        }


        /// <summary>
        /// 单元格自动换行
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        public void AutoWrapText(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell)
        {
            CurSheet.get_Range(objStartCell, objEndCell).WrapText = true;
        }

        /// <summary>
        /// 设置整个连续区域的字体颜色
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        /// <param name="clrColor">颜色</param>
        public void SetColor(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell, System.Drawing.Color clrColor)
        {
            CurSheet.get_Range(objStartCell, objEndCell).Font.Color = System.Drawing.ColorTranslator.ToOle(clrColor);
        }

        /// <summary>
        /// 设置整个连续区域的单元格背景色
        /// </summary>
        /// <param name="CurSheet"></param>
        /// <param name="objStartCell"></param>
        /// <param name="objEndCell"></param>
        /// <param name="clrColor"></param>
        public void SetBgColor(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell, System.Drawing.Color clrColor)
        {
            CurSheet.get_Range(objStartCell, objEndCell).Interior.Color = System.Drawing.ColorTranslator.ToOle(clrColor);
        }

        /// <summary>
        /// 设置连续区域的字体名称
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        /// <param name="fontname">字体名称 隶书、仿宋_GB2312等</param>
        public void SetFontName(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell, string fontname)
        {
            CurSheet.get_Range(objStartCell, objEndCell).Font.Name = fontname;
        }

        /// <summary>
        /// 设置连续区域的字体为黑体
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        public void SetBold(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell)
        {
            CurSheet.get_Range(objStartCell, objEndCell).Font.Bold = true;
        }


        /// <summary>
        /// 设置连续区域的边框：上下左右都为黑色连续边框
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        public void SetBorderAll(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell)
        {
            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            CurSheet.get_Range(objStartCell, objEndCell).Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

        }

        /// <summary>
        /// 设置连续区域水平居中
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        public void SetHAlignCenter(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell)
        {
            CurSheet.get_Range(objStartCell, objEndCell).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }

        /// <summary>
        /// 设置连续区域水平居左
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        public void SetHAlignLeft(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell)
        {
            CurSheet.get_Range(objStartCell, objEndCell).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
        }

        /// <summary>
        /// 设置连续区域水平居右
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        public void SetHAlignRight(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell)
        {
            CurSheet.get_Range(objStartCell, objEndCell).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
        }


        /// <summary>
        /// 设置连续区域的显示格式
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        /// <param name="strNF">如"#,##0.00"的显示格式</param>
        public void SetNumberFormat(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell, string strNF)
        {
            CurSheet.get_Range(objStartCell, objEndCell).NumberFormat = strNF;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="strColID">列标识，如A代表第一列</param>
        /// <param name="dblWidth">宽度</param>
        public void SetColumnWidth(Microsoft.Office.Interop.Excel._Worksheet CurSheet, string strColID, double dblWidth)
        {
            ((Microsoft.Office.Interop.Excel.Range)CurSheet.Columns.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, CurSheet.Columns, new object[] { (strColID + ":" + strColID).ToString() })).ColumnWidth = dblWidth;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        /// <param name="dblWidth">宽度</param>
        public void SetColumnWidth(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell, double dblWidth)
        {
            CurSheet.get_Range(objStartCell, objEndCell).ColumnWidth = dblWidth;
        }


        /// <summary>
        /// 设置行高
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objStartCell">开始单元格</param>
        /// <param name="objEndCell">结束单元格</param>
        /// <param name="dblHeight">行高</param>
        public void SetRowHeight(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objStartCell, object objEndCell, double dblHeight)
        {
            CurSheet.get_Range(objStartCell, objEndCell).RowHeight = dblHeight;
        }


        /// <summary>
        /// 为单元格添加超级链接
        /// </summary>
        /// <param name="CurSheet">Worksheet</param>
        /// <param name="objCell">单元格</param>
        /// <param name="strAddress">链接地址</param>
        /// <param name="strTip">屏幕提示</param>
        /// <param name="strText">链接文本</param>
        public void AddHyperLink(Microsoft.Office.Interop.Excel._Worksheet CurSheet, object objCell, string strAddress, string strTip, string strText)
        {
            CurSheet.Hyperlinks.Add(CurSheet.get_Range(objCell, objCell), strAddress, mValue, strTip, strText);
        }

        /// <summary>
        /// 另存为xls文件
        /// </summary>
        /// <param name="CurBook">Workbook</param>
        /// <param name="strFilePath">文件路径</param>
        public void Save(Microsoft.Office.Interop.Excel._Workbook CurBook, string strFilePath)
        {
            CurBook.SaveCopyAs(strFilePath);
        }

        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="CurBook">Workbook</param>
        /// <param name="strFilePath">文件路径</param>
        public void SaveAs(Microsoft.Office.Interop.Excel._Workbook CurBook, string strFilePath)
        {
            CurBook.SaveAs(strFilePath, mValue, mValue, mValue, mValue, mValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, mValue, mValue, mValue, mValue, mValue);
        }

        /// <summary>
        /// 另存为html文件
        /// </summary>
        /// <param name="CurBook">Workbook</param>
        /// <param name="strFilePath">文件路径</param>
        public void SaveHtml(Microsoft.Office.Interop.Excel._Workbook CurBook, string strFilePath)
        {
            CurBook.SaveAs(strFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml, mValue, mValue, mValue, mValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, mValue, mValue, mValue, mValue, mValue);
        }


        /// <summary>
        /// 释放内存
        /// </summary>
        public void Dispose(Microsoft.Office.Interop.Excel._Worksheet CurSheet, Microsoft.Office.Interop.Excel._Workbook CurBook, Microsoft.Office.Interop.Excel._Application CurExcel)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(CurSheet);
                CurSheet = null;
                CurBook.Close(false, mValue, mValue);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(CurBook);
                CurBook = null;

                CurExcel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(CurExcel);
                CurExcel = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (System.Exception ex)
            {
                // HttpContext.Current.Response.Write("在释放Excel内存空间时发生了一个错误:" + ex);
            }
            finally
            {
                foreach (System.Diagnostics.Process pro in System.Diagnostics.Process.GetProcessesByName("Excel"))
                    //if (pro.StartTime < DateTime.Now)
                    pro.Kill();
            }
            System.GC.SuppressFinalize(this);

        }
    }
}
