//using System;
//using System.Data;
//using System.IO;
//using System.Text;

//namespace Lau.Net.Utils
//{
//   public class ExcelUtil
//    {
//       /// <summary>
//        /// 大批量数据导出到excel(使用StreamWriter操作）
//       /// 问题1、所以导出的数据都会弹出一个提示框，“您尝试打开的文件的格式与文件扩展名指定格式不一致，打开文件前请验证文件没有损坏且来源可信”的对话框
//       /// 问题2、当修改里面的内容时只能保存一个副本操作及其不方便。
//       /// </summary>
//       /// <param name="path">保存路径</param>
//       /// <param name="sourceTable">源数据表</param>
//       /// <param name="encoding">编码，默认为gb2312</param>
//       public static void DataTableToExcel(string path,DataTable sourceTable,Encoding encoding=null)
//       {
//           if (encoding == null) Encoding.GetEncoding("gb2312");
//               StreamWriter sw = new StreamWriter(path, false, encoding);
//               StringBuilder sb = new StringBuilder();
//           try
//           {
//               for (int i = 0; i < sourceTable.Columns.Count; i++)
//               {
//                   sb.Append(sourceTable.Columns[i].ColumnName.ToString() + "\t");
//               }
//               sb.Append(Environment.NewLine);

//               for (int i = 0; i < sourceTable.Rows.Count; i++)
//               {
//                   //rowRead++;
//                   //percent = ((float)(100 * rowRead)) / totalCount;
//                   //Pbar.Maximum = (int)totalCount;
//                   //Pbar.Value = (int)rowRead;
//                   //lblTip.Text = "正在写入[" + percent.ToString("0.00") + "%]...的数据";
//                   //System.Windows.Forms.Application.DoEvents();
//                   for (int j = 0; j < sourceTable.Columns.Count; j++)
//                   {
//                       sb.Append(sourceTable.Rows[i][j].ToString() + "\t");
//                   }
//                   sb.Append(Environment.NewLine);
//               }
//               sw.Write(sb.ToString());
//               sw.Flush();
//               sw.Close();
//               MsgBox.ShowInformation("导出EXCEL成功");
//           }
//           catch (Exception ex)
//           {
//               MsgBox.ShowError("导出EXCEL失败:" + ex.Message);
//           }


//       }

//       /// <summary>
//       /// 把一个数据集中的数据导出到Excel文件中(XML格式操作)
//       /// 问题1、所以导出的数据都会弹出一个提示框，“您尝试打开的文件的格式与文件扩展名指定格式不一致，打开文件前请验证文件没有损坏且来源可信”的对话框
//       /// 问题2、当修改里面的内容时只能保存一个副本操作及其不方便。
//       /// </summary>
//       /// <param name="dataSet">DataSet数据</param> 
//       public static void DataSetToExcel(DataSet dataSet)
//       {
//           string fileName =FileDialogUtil.SaveExcel();
//           if (string.IsNullOrEmpty(fileName)) return;
//           #region Excel格式内容
//           var excelDoc = new StreamWriter(fileName);
//           const string startExcelXML = "<xml version>\r\n<Workbook " +
//                 "xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\r\n" +
//                 " xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n " +
//                 "xmlns:x=\"urn:schemas-    microsoft-com:office:" +
//                 "excel\"\r\n xmlns:ss=\"urn:schemas-microsoft-com:" +
//                 "office:spreadsheet\">\r\n <Styles>\r\n " +
//                 "<Style ss:ID=\"Default\" ss:Name=\"Normal\">\r\n " +
//                 "<Alignment ss:Vertical=\"Bottom\"/>\r\n <Borders/>" +
//                 "\r\n <Font/>\r\n <Interior/>\r\n <NumberFormat/>" +
//                 "\r\n <Protection/>\r\n </Style>\r\n " +
//                 "<Style ss:ID=\"BoldColumn\">\r\n <Font " +
//                 "x:Family=\"Swiss\" ss:Bold=\"1\"/>\r\n </Style>\r\n " +
//                 "<Style     ss:ID=\"StringLiteral\">\r\n <NumberFormat" +
//                 " ss:Format=\"@\"/>\r\n </Style>\r\n <Style " +
//                 "ss:ID=\"Decimal\">\r\n <NumberFormat " +
//                 "ss:Format=\"#,##0.###\"/>\r\n </Style>\r\n " +
//                 "<Style ss:ID=\"Integer\">\r\n <NumberFormat " +
//                 "ss:Format=\"0\"/>\r\n </Style>\r\n <Style " +
//                 "ss:ID=\"DateLiteral\">\r\n <NumberFormat " +
//                 "ss:Format=\"yyyy-mm-dd;@\"/>\r\n </Style>\r\n " +
//                 "</Styles>\r\n ";
//           const string endExcelXML = "</Workbook>";
//           #endregion

//           int sheetCount = 1;
//           excelDoc.Write(startExcelXML);
//           for (int i = 0; i < dataSet.Tables.Count; i++)
//           {
//               int rowCount = 0;
//               DataTable dt = dataSet.Tables[i];

//               excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
//               excelDoc.Write("<Table>");
//               excelDoc.Write("<Row>");
//               for (int x = 0; x < dt.Columns.Count; x++)
//               {
//                   excelDoc.Write("<Cell ss:StyleID=\"BoldColumn\"><Data ss:Type=\"String\">");
//                   excelDoc.Write(dataSet.Tables[0].Columns[x].ColumnName);
//                   excelDoc.Write("</Data></Cell>");
//               }
//               excelDoc.Write("</Row>");
//               foreach (DataRow x in dt.Rows)
//               {
//                   rowCount++;
//                   //if the number of rows is > 64000 create a new page to continue output

//                   if (rowCount == 64000)
//                   {
//                       rowCount = 0;
//                       sheetCount++;
//                       excelDoc.Write("</Table>");
//                       excelDoc.Write(" </Worksheet>");
//                       excelDoc.Write("<Worksheet ss:Name=\"Sheet" + sheetCount + "\">");
//                       excelDoc.Write("<Table>");
//                   }
//                   excelDoc.Write("<Row>"); //ID=" + rowCount + "

//                   for (int y = 0; y < dataSet.Tables[0].Columns.Count; y++)
//                   {
//                       Type rowType = x[y].GetType();
//                       #region 根据不同数据类型生成内容
//                       switch (rowType.ToString())
//                       {
//                           case "System.String":
//                               string XMLstring = x[y].ToString();
//                               XMLstring = XMLstring.Trim();
//                               XMLstring = XMLstring.Replace("&", "&");
//                               XMLstring = XMLstring.Replace(">", ">");
//                               XMLstring = XMLstring.Replace("<", "<");
//                               excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
//                                              "<Data ss:Type=\"String\">");
//                               excelDoc.Write(XMLstring);
//                               excelDoc.Write("</Data></Cell>");
//                               break;
//                           case "System.DateTime":
//                               //Excel has a specific Date Format of YYYY-MM-DD followed by
//                               //the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
//                               //The Following Code puts the date stored in XMLDate
//                               //to the format above

//                               var XMLDate = (DateTime)x[y];

//                               string XMLDatetoString = XMLDate.Year +
//                                                        "-" +
//                                                        (XMLDate.Month < 10
//                                                             ? "0" +
//                                                               XMLDate.Month
//                                                             : XMLDate.Month.ToString()) +
//                                                        "-" +
//                                                        (XMLDate.Day < 10
//                                                             ? "0" +
//                                                               XMLDate.Day
//                                                             : XMLDate.Day.ToString()) +
//                                                        "T" +
//                                                        (XMLDate.Hour < 10
//                                                             ? "0" +
//                                                               XMLDate.Hour
//                                                             : XMLDate.Hour.ToString()) +
//                                                        ":" +
//                                                        (XMLDate.Minute < 10
//                                                             ? "0" +
//                                                               XMLDate.Minute
//                                                             : XMLDate.Minute.ToString()) +
//                                                        ":" +
//                                                        (XMLDate.Second < 10
//                                                             ? "0" +
//                                                               XMLDate.Second
//                                                             : XMLDate.Second.ToString()) +
//                                                        ".000";
//                               excelDoc.Write("<Cell ss:StyleID=\"DateLiteral\">" +
//                                            "<Data ss:Type=\"DateTime\">");
//                               excelDoc.Write(XMLDatetoString);
//                               excelDoc.Write("</Data></Cell>");
//                               break;
//                           case "System.Boolean":
//                               excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
//                                           "<Data ss:Type=\"String\">");
//                               excelDoc.Write(x[y].ToString());
//                               excelDoc.Write("</Data></Cell>");
//                               break;
//                           case "System.Int16":
//                           case "System.Int32":
//                           case "System.Int64":
//                           case "System.Byte":
//                               excelDoc.Write("<Cell ss:StyleID=\"Integer\">" +
//                                       "<Data ss:Type=\"Number\">");
//                               excelDoc.Write(x[y].ToString());
//                               excelDoc.Write("</Data></Cell>");
//                               break;
//                           case "System.Decimal":
//                           case "System.Double":
//                               excelDoc.Write("<Cell ss:StyleID=\"Decimal\">" +
//                                     "<Data ss:Type=\"Number\">");
//                               excelDoc.Write(x[y].ToString());
//                               excelDoc.Write("</Data></Cell>");
//                               break;
//                           case "System.DBNull":
//                               excelDoc.Write("<Cell ss:StyleID=\"StringLiteral\">" +
//                                     "<Data ss:Type=\"String\">");
//                               excelDoc.Write("");
//                               excelDoc.Write("</Data></Cell>");
//                               break;
//                           default:
//                               throw (new Exception(rowType.ToString() + " not handled."));
//                       }
//                       #endregion
//                   }
//                   excelDoc.Write("</Row>");
//               }
//               excelDoc.Write("</Table>");
//               excelDoc.Write(" </Worksheet>");

//               sheetCount++;
//           }

//           excelDoc.Write(endExcelXML);
//           excelDoc.Close();
//       }


//        #region DataGridView导入Excel，需调用Microsoft.Office.Interop.Excel
//        //public static void ExportExcel(DataGridView dgv)
//       //{
//       //    SaveFileDialog sfd = new SaveFileDialog();
//       //    sfd.Filter = "Excel 文件 (*.xls)|*.xls|Excel 文件(*.xlsx)|*.xlsx";
//       //    sfd.FilterIndex = 0;
//       //    sfd.RestoreDirectory = true;
//       //    sfd.CreatePrompt = true;
//       //    sfd.Title = "請選擇文檔保存路徑";
//       //    sfd.ShowDialog();
//       //    string strName = sfd.FileName;

//       //    if (strName.Length != 0)
//       //    {
//       //        System.Reflection.Missing miss = System.Reflection.Missing.Value;
//       //        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
//       //        excel.Application.Workbooks.Add(true);
//       //        excel.Visible = false;  //若是true，則在導出的時候會顯示EXcel介面。
//       //        if (excel == null)
//       //        {
//       //            MessageBox.Show("Excel無法啟動", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
//       //            return;
//       //        }
//       //        Workbooks books = (Microsoft.Office.Interop.Excel.Workbooks)excel.Workbooks;
//       //        Workbook book = (Microsoft.Office.Interop.Excel.Workbook)books.Add(miss);
//       //        Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet;
//       //        sheet.Name = "sheet1";

//       //        int n = 0;
//       //        //生成列名稱
//       //        for (int i = 0; i < dgv.ColumnCount; i++)
//       //        {
//       //            if (dgv.Columns[i].Visible == false) continue;
//       //            n++;
//       //            excel.Cells[1, n] = dgv.Columns[i].HeaderText.ToString();
//       //        }
//       //        //填充數據
//       //        for (int i = 0; i < dgv.RowCount; i++)
//       //        {
//       //            n = 0;
//       //            for (int j = 0; j < dgv.ColumnCount; j++)
//       //            {
//       //                if (dgv[j, i].Visible == false) continue;
//       //                n++;
//       //                if (dgv[j, i].Value.GetType() == typeof(string))
//       //                {
//       //                    excel.Cells[i + 2, n] = "'" + dgv[j, i].Value.ToString();
//       //                }
//       //                else
//       //                {
//       //                    excel.Cells[i + 2, n] = dgv[j, i].Value.ToString();
//       //                }
//       //            }
//       //        }
//       //        try
//       //        {
//       //            sheet.SaveAs(strName, miss, miss, miss, miss, miss, XlSaveAsAccessMode.xlNoChange, miss, miss, miss);
//       //        }
//       //        catch (Exception ex)
//       //        {
//       //            //throw new Exception(ex.Message, ex);
//       //            MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
//       //            return;
//       //        }
//       //        book.Close(false, miss, miss);
//       //        books.Close();
//       //        excel.Quit();
//       //        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
//       //        System.Runtime.InteropServices.Marshal.ReleaseComObject(book);
//       //        System.Runtime.InteropServices.Marshal.ReleaseComObject(books);
//       //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
//       //        GC.Collect();
//       //        MessageBox.Show("Excel導出成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
//       //        System.Diagnostics.Process.Start(strName);
//       //    }

//       //}
//        #endregion

//        #region  VB写的Excel报表
//                //       Dim exapp As New Excel.Application
//                //Dim exbook As Excel.Workbook
//                //Dim exsheet As Excel.Worksheet
//                //exapp = CreateObject("Excel.Application")
//                //exbook = exapp.Workbooks.Add
//                //exapp.Visible = True
//                //exapp.Sheets.Add()
//                //exsheet = exapp.Worksheets(1)
//                //exsheet.Name = "asdf"
//                //exsheet.Cells.Font.Size = 10
//                //exsheet.Shapes.AddPicture("C:\LF_Logo.JPG", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 10, 10, 90, 60)
//                //exsheet.Range(exsheet.Cells(1, 1), exsheet.Cells(5, 9)).Merge()
//                //exsheet.Range("A1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
//                //exsheet.Range("A1").VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
//                //exsheet.Range("A1").Font.Size = 9
//                //exsheet.Cells(1, 1) = "聯豐表殼廠有限公司" + vbCrLf + "LUEN FUNG WATCH CASE FACTORY LIMITED" + vbCrLf + "香港新界葵芳與芳路223號新都會廣場一期28樓2801-06&2852室" + vbCrLf + "Tel:(86) 852 2480 6818  Fax:(86) 852 2487 9780" + vbCrLf + "深圳市寶安區民治街道上芬社區中環路70" + vbCrLf + "Tel:(86) 755 2774 8020  Fax:(86) 755 2774 9753"
//                //exsheet.Cells(6, 5) = "Quotation"
//                //exsheet.Range("E6").Font.Bold = True
//                //exsheet.Range("E6").Font.Size = 16
//                //exsheet.Columns(5).columnwidth = 16
//                //exsheet.Columns(8).columnwidth = 11
//                //exsheet.Range(exsheet.Cells(8, 2), exsheet.Cells(9, 5)).Merge()
//                //exsheet.Range(exsheet.Cells(8, 2), exsheet.Cells(9, 5)).ShrinkToFit = True
//                //exsheet.Range(exsheet.Cells(8, 2), exsheet.Cells(9, 5)).WrapText = True
//                //exsheet.Range(exsheet.Cells(7, 1), exsheet.Cells(12, 1)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
//                //exsheet.Range(exsheet.Cells(7, 7), exsheet.Cells(12, 7)).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
//                //exsheet.Range(exsheet.Cells(7, 8), exsheet.Cells(12, 8)).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
//                //exsheet.Cells(7, 1) = "Company:"
//                //exsheet.Cells(8, 1) = "Address:"
//                //exsheet.Cells(10, 1) = "Attention:"
//                //exsheet.Cells(11, 1) = "Project:"
//                //exsheet.Cells(12, 1) = "Handled By:"
//                //exsheet.Cells(7, 8) = "Ref:"
//                //exsheet.Cells(8, 8) = "Date:"
//                //exsheet.Cells(10, 8) = "Tel No:"
//                //exsheet.Cells(11, 8) = "Fax No:"
//                //exsheet.Cells(12, 8) = "cc:"

//                //exsheet.Cells(7, 2) = corpInfo.Company
//                //exsheet.Cells(8, 2) = corpInfo.Address
//                //exsheet.Cells(10, 2) = corpInfo.Attention
//                //exsheet.Cells(11, 2) = corpInfo.Project
//                //exsheet.Cells(12, 2) = corpInfo.HandledBy
//                //exsheet.Cells(7, 9) = corpInfo.Ref
//                //exsheet.Cells(8, 9) = Format(CDate(corpInfo.CreateDate), "yyyy/MM/dd")
//                //exsheet.Cells(10, 9) = corpInfo.TelNo
//                //exsheet.Cells(11, 9) = corpInfo.FaxNo
//                //exsheet.Cells(12, 9) = corpInfo.cc
//                //exsheet.Range(exsheet.Cells(12, 1), exsheet.Cells(12, 9)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2

//                //exsheet.Range("A13").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
//                //exsheet.Range("C13").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
//                //exsheet.Range("E13").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
//                //exsheet.Range("G13").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
//                //exsheet.Range("I13").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
//                //exsheet.Cells(13, 1) = "Item"
//                //exsheet.Cells(13, 3) = "Description"
//                //exsheet.Cells(13, 5) = "Quantity"
//                //exsheet.Cells(13, 7) = "New Price"
//                //exsheet.Cells(13, 9) = "Total(" + corpInfo.Currency + ")$"


//                //Dim alist As New List(Of QuotedPrice_DetailInfo)
//                //alist = qpdCont.QuotedPrice_DetailA_GetListReport(strAutoID, Nothing, Nothing, True, True)
//                //exsheet.Cells(14, 1) = "A    Case  Swiss Model:" + corpInfo.StyleNo
//                //exsheet.Range("A14").Font.Bold = True
//                //exsheet.Shapes.AddPicture("C:\LF_Logo.JPG", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 210, 550, alist.Count * 14.6)
//                //Dim rowNum As Integer = 15
//                //For i As Integer = 0 To alist.Count - 1
//                //    exsheet.Cells(rowNum, 3) = "~" + alist(i).Endescription
//                //    exsheet.Cells(rowNum, 5) = alist(i).Quantity + " PCS"
//                //    exsheet.Cells(rowNum, 7) = alist(i).NetPrice + " $"
//                //    exsheet.Cells(rowNum, 9) = alist(i).TotalPrice + " $"
//                //    rowNum = rowNum + 1
//                //Next

//                //Dim blist As New List(Of QuotedPrice_DetailInfo)
//                //blist = qpdCont.QuotedPrice_DetailB_GetList(strAutoID, Nothing, Nothing, True)
//                //exsheet.Cells(rowNum, 1) = "B    Additional Cost"
//                //exsheet.Range("A" + rowNum.ToString).Font.Bold = True
//                //rowNum = rowNum + 1
//                //For i As Integer = 0 To blist.Count - 1
//                //    exsheet.Cells(rowNum, 3) = "~" + blist(i).Endescription
//                //    exsheet.Cells(rowNum, 5) = blist(i).Quantity + " PCS"
//                //    exsheet.Cells(rowNum, 7) = blist(i).NetPrice + " $"
//                //    exsheet.Cells(rowNum, 9) = blist(i).TotalPrice + " $"
//                //    rowNum = rowNum + 1
//                //Next

//                //Dim clist As New List(Of QuotedPrice_DetailInfo)
//                //clist = qpdCont.QuotedPrice_DetailC_GetList(strAutoID, Nothing, Nothing, True)
//                //exsheet.Cells(rowNum, 1) = "C    Plating Cost"
//                //exsheet.Range("A" + rowNum.ToString).Font.Bold = True
//                //rowNum = rowNum + 1
//                //For i As Integer = 0 To clist.Count - 1
//                //    exsheet.Cells(rowNum, 3) = "~" + clist(i).Endescription
//                //    exsheet.Cells(rowNum, 5) = clist(i).Quantity + " PCS"
//                //    exsheet.Cells(rowNum, 7) = clist(i).NetPrice + " $"
//                //    exsheet.Cells(rowNum, 9) = clist(i).TotalPrice + " $"
//                //    rowNum = rowNum + 1
//                //Next
//                //exsheet.Range(exsheet.Cells(rowNum - 1, 1), exsheet.Cells(rowNum - 1, 9)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
//                //exsheet.Range("I" + rowNum.ToString).Font.Bold = True
//                //exsheet.Range("I" + rowNum.ToString).Formula = "=Sum(I14:I" + (rowNum - 1).ToString
//                //exsheet.Cells(rowNum, 8) = "Total For Each Set:" + corpInfo.Currency + "$"
//                //exsheet.Range(exsheet.Cells(rowNum, 8), exsheet.Cells(rowNum, 9)).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
//                //rowNum = rowNum + 1

//                //exsheet.Range("A" + rowNum.ToString).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
//                //exsheet.Range("A" + rowNum.ToString).Font.Bold = True
//                //exsheet.Range("A" + rowNum.ToString).Value = "Terms&Conditions"
//                //rowNum = rowNum + 1
//                //exsheet.Range(exsheet.Cells(rowNum, 2), exsheet.Cells(rowNum, 9)).Merge()
//                //'exsheet.Range(exsheet.Rows(rowNum), exsheet.Rows(rowNum)).WrapText = True
//                //' exsheet.Range("A" + rowNum.ToString).LocationInTable()
//                //Dim company_subList As New List(Of Company_SubInfo)
//                //company_subList = comCont.Company_Sub_GetList(strCompanyID, Nothing, True)
//                //exsheet.Range(exsheet.Cells(rowNum, 1), exsheet.Cells(rowNum + company_subList.Count - 1, 1)).Font.Bold = True
//                //For i As Integer = 0 To company_subList.Count - 1
//                //    exsheet.Cells(rowNum, 1) = company_subList(i).Title
//                //    exsheet.Cells(rowNum, 2) = company_subList(i).Neirong
//                //    rowNum = rowNum + 1
//                //Next
//                //exsheet.Range(exsheet.Cells(rowNum - company_subList.Count, 1), exsheet.Cells(rowNum, 1)).Font.Size = 9

//                //rowNum = rowNum + 1      '空一行
//                //exsheet.Cells(rowNum, 1) = "For and on behalf of"
//                //exsheet.Cells(rowNum + 1, 1) = "LUEN FUNG WATCH CASE FACTORY LIMITED"
//                //exsheet.Cells(rowNum + 1, 8) = "Accepted By :"
//                //exsheet.Range("A" + (rowNum + 2).ToString).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
//                //exsheet.Range("H" + (rowNum + 2).ToString).Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
//                //exsheet.Cells(rowNum + 3, 1) = "Signature & Chop"
//                //exsheet.Cells(rowNum + 3, 8) = "Signature & Chop"
//                //exsheet.Cells(rowNum + 4, 1) = "Date:" + Now.Date.ToString("yyyy/MM/dd")
//                //exsheet.Cells(rowNum + 4, 8) = "Date:"

//                //exsheet.PageSetup.LeftMargin = 0.25
//                //exsheet.PageSetup.RightMargin = 0.25
//        #endregion
//    }
//}
