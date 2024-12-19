/* Version 18-05-2016 cambio el 28-10-2016*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
//using System.Collections.Concurrent;

namespace Interface.ExcelTools
{
    internal class ExcelInterface : IDisposable
    {
        public Excel.Application app;
        public Excel.Range range;
        public Excel.Workbook workbook;
        public Excel.Worksheet worksheet;
        public Excel.Range rngmem;
        public Excel.Workbook bookmem;
        public Excel.Worksheet sheetmem;
        private object misValue = System.Reflection.Missing.Value;

        #region Constructor and Dispose Method

        /// <summary>
        /// Limpia variables
        /// </summary>
        public void ClearVars()
        {
            this.app = null;
            this.workbook = null;
            this.bookmem = null;
            this.worksheet = null;
            this.sheetmem = null;
            this.range = null;
            this.rngmem = null;
        }

        /// <summary>
        /// Libera variables y desvincula el programa
        /// </summary>
        /// <returns></returns>
        private bool ReleaseObjects()
        {
            try
            {
                if (this.app != null) releaseObject(this.app);
                if (this.workbook != null) releaseObject(this.workbook);
                if (this.bookmem != null) releaseObject(this.bookmem);
                if (this.worksheet != null) releaseObject(this.worksheet);
                if (this.sheetmem != null) releaseObject(this.sheetmem);
                if (this.range != null) releaseObject(this.range);
                if (this.rngmem != null) releaseObject(this.rngmem);
                this.ClearVars();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public ExcelInterface()
        {
            this.ClearVars();
        }

        public void Dispose()
        {
            this.ReleaseObjects();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Releases the object.
        /// </summary>
        /// <param name="obj">The object.</param>
        /// <exception cref="System.Exception">Unable to release the Object + ex.Message</exception>
        /// Method created by Silvonei: 2013-02-20.
        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch {}
            //catch (Exception ex)
            //{
                //throw new Exception("Unable to release the Object " + ex.Message);
            //}
            finally
            {
                GC.Collect();
                obj = null;
            }
        }

        #endregion

        #region ConnectionWithXLS
        /// <summary>
        /// Abre Excel, y abre el archivo fileName en modo solo lectura
        /// </summary>
        /// <param name="fileName">Dirección completa del archivo excel</param>
        /// <returns></returns>
        public bool OpenFile(string fileName)
        {
            if (!System.IO.File.Exists(fileName))
                return false;
            try
            {
                if (this.app == null) 
                    this.app = new Excel.Application();
                this.workbook = this.app.Workbooks.Open(fileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                if (this.workbook == null) return false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Abre Excel, y abre el archivo fileName en modo escritura
        /// </summary>
        /// <param name="fileName">Dirección completa del archivo excel</param>
        /// <returns></returns>
        public bool OpenFileToEdit(string fileName)
        {
            if (!System.IO.File.Exists(fileName))
                return false;
            try
            {
                if (this.app == null)
                    this.app = new Excel.Application();
                this.workbook = this.app.Workbooks.Open(fileName);
                if (this.workbook == null) return false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Cierra el documento, cierra el Excel
        /// </summary>
        /// <returns></returns>
        public bool CloseXLS()
        {
            try
            {
                this.workbook.Close(false, false, true);
                this.app.Quit();
            }
            catch
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Libera variables y desvincula el programa
        /// </summary>
        /// <returns></returns>
        public bool ReleaseWithoutCloseXLS()
        {
            if (ReleaseObjects())
                return true;
            else
                return false;
        }

        /// <summary>
        /// Abre Excel, crea un documento nuevo
        /// </summary>
        /// <returns></returns>
        public bool CreateNewFile()
        {
            try
            {
                if (this.app == null) 
                    this.app = new Excel.Application();
                this.workbook = app.Workbooks.Add(Type.Missing);
                if (this.workbook == null) return false;
                return true;
            }
            catch (Exception)
            {
                //throw new Exception(String.Format("Error al crear nuevo documento en Excel.{0}{1}", Environment.NewLine, ex.Message));
                return false;
            }
        }

        /// <summary>
        /// Conecta con un Excel abierto
        /// </summary>
        /// <returns></returns>
        public bool ConnectWithApp()
        {
            //this.app = null;
            try
            {
                //Checks to see if excel is opened
                this.app = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception) //Excel not open
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// Proceso para buscar un archivo y conectarlo
        /// </summary>
        /// <param name="strFileFullName"></param>
        /// <returns></returns>
        public bool FindDocument(string strFileFullName)
        {
            //this.app = null;
            this.workbook = null;
            if (!System.IO.File.Exists(strFileFullName))
                return false;
            string strFileName = System.IO.Path.GetFileName(strFileFullName);

            if (this.ConnectWithApp())
            {
                if (this.SelectDocument(strFileName))
                    return true;
            }
            if (this.workbook == null)
            {
                this.OpenFileToEdit(strFileFullName);
            }
            if (this.workbook == null)
                return false;
            return true;
        }

        /// <summary>
        /// Elegir documento entre tantos Excels abiertos
        /// </summary>
        /// <param name="strFileName"></param>
        /// <returns></returns>
        public bool SelectDocument(string strFileName)
        {
            if (this.app == null) return false;
            foreach (Excel.Workbook wkb in this.app.Application.Workbooks)
            {
                if (wkb.Name == strFileName)
                {
                    this.workbook = wkb;
                    if (this.workbook != null)
                        return true;
                }
            }
            return false;
        }

        public bool GetOpenedFile()
        {
            app = null;
            try
            {
                //Checks to see if excel is opened
                app = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
                workbook = app.Workbooks.Add(Type.Missing);
            }
            catch (Exception)//Excel not open
            {
                return false;
            }
            finally
            {
                //Release the temp if in use
                if (null != app) { Marshal.FinalReleaseComObject(app); }
                app = null;
            }
            return true;
        }

        #endregion

        /// <summary>
        /// Elegir hoja entre tantos Documentos abiertos
        /// </summary>
        /// <param name="shtname"></param>
        /// <returns></returns>
        public bool SelectSheet(string shtname)
        {
            if (this.workbook == null) return false;
            foreach (Excel.Worksheet wksht in workbook.Worksheets)
            {
                if (wksht.Name == shtname)
                {
                    this.worksheet = wksht;
                    if (this.worksheet != null)
                        return true;
                }
            }
            return false;
        }

        public bool CopySheet(string shtname, string newname)
        {
            try
            {
                this.SelectSheet(shtname);
                if (this.worksheet == null) return false;
                this.worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                this.worksheet.Select();
                this.worksheet.Copy(this.workbook.Worksheets[this.workbook.Worksheets.Count]);
                this.worksheet.Select();
                this.worksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                this.SelectSheet(shtname + " (2)");
                this.worksheet.Select();
                this.worksheet.Name = newname;
            }
            catch
            {
                return false;
            }
            return true;
        }

        public bool TrySaveAs(string strFileName)
        {
            if (this.workbook == null) return false;
            try
            {
                workbook.SaveAs(strFileName, Type.Missing); //56 o -4143
                return true;
            }
            catch { }
            return false;
        }

        public void ShowExcel()
        {
            if (this.app == null) return;
            this.app.WindowState = Excel.XlWindowState.xlMaximized;
            this.app.Visible = true;
        }

        public void HideExcel()
        {
            if (this.app == null) return;
            this.app.WindowState = Excel.XlWindowState.xlMinimized;
            this.app.Visible = false;
        }

        public object[,] GetSheetInfo(int sheetnumber)
        {
            try
            {
                Excel.Worksheet xlSheet = (Excel.Worksheet)this.workbook.Worksheets.get_Item((int)sheetnumber);
                Excel.Range xlRange = xlSheet.UsedRange;

                //string strExcelName = xlBook.Name;
                //strExcelName = strExcelName.Substring(0, strExcelName.IndexOf('.'));

                object[,] exData = xlRange.Text as object[,];
                releaseObject(xlSheet);
                xlRange = null;
                xlSheet = null;
                return exData;

            }
            catch
            {
                return null;
            }
        }

        public string GetSheetName(int sheetnumber)
        {
            string shtname = string.Empty;
            try
            {
                Excel.Worksheet xlSheet = (Excel.Worksheet)this.workbook.Worksheets.get_Item((int)sheetnumber);
                shtname = xlSheet.Name;
                releaseObject(xlSheet);
                xlSheet = null;
                return shtname;
            }
            catch
            {
                return shtname;
            }
        }

        public int GetLastUsedRow()
        {
            Excel.Range lastCell = this.worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastRow = lastCell.Row;
            int lastCol = lastCell.Column;
            lastCell = null;
            return lastRow;
        }

        public void CreateSheet(string shtname)
        {
            try
            {
                if (shtname == "1")
                    this.worksheet = (Excel.Worksheet)this.workbook.Sheets[1];
                else
                {
                    this.worksheet = (Excel.Worksheet)this.workbook.Sheets.Add(System.Reflection.Missing.Value,
                            workbook.Worksheets[workbook.Worksheets.Count],
                            System.Reflection.Missing.Value,
                            System.Reflection.Missing.Value);
                }
                this.worksheet.Name = shtname;
            }
            catch (Exception ex)
            {
                throw new Exception(String.Format("Error al crear nuevo documento en Excel.{0}{1}", Environment.NewLine, ex.Message));
            }
        }

        public Boolean DeleteSheet(string shtname)
        {
            try
            {
                if (this.workbook == null) return false;
                foreach (Excel.Worksheet wksht in workbook.Worksheets)
                {
                    if (wksht.Name == shtname)
                    {
                        this.worksheet = wksht;
                        this.worksheet.Delete();
                        if (this.worksheet != null)
                            return true;
                    }
                }
                return false;

                //Excel.Worksheet myWorksheet = (Worksheet) workbook.Worksheets[1];
                //myWorksheet.Delete();
            }
            catch (Exception ex)
            {
                throw new Exception(String.Format("Error al borrar hoja en documento en Excel.{0}{1}", Environment.NewLine, ex.Message));
                //return false;
            }
        }


        #region Funciones_Independientes
        public string GetData(object data)
        {
            if (data == null) return string.Empty;
            return data.ToString();
        }

        public string GetRowAndColValue(int row, int argcol)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string value = "";
            int col = argcol - 1;

            if (col >= letters.Length)
                value += letters[col / letters.Length - 1];

            value += letters[col % letters.Length];

            return (value + row.ToString()).ToString();
        }
        #endregion

        public void SelectRange(string cellrange)
        {
            if (this.worksheet == null) return;
            if (cellrange.Contains('*'))
            {
                Excel.Range last = this.worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                this.range = (Excel.Range)this.worksheet.get_Range(cellrange.Replace("*", ""), last);
            }
            else if (cellrange.Contains(':'))
            {
                //const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                //if (cellrange.Split(':')[0]. > 0)
                //this.range = (Excel.Range)worksheet.Columns[colrange];
                //this.range = (Excel.Range)worksheet.Rows[rowrange];

            }
            else if (cellrange.Contains(','))
                this.range = (Excel.Range)this.worksheet.get_Range(cellrange.Split(',')[0], cellrange.Split(',')[1]);
            else if (cellrange.Contains('-'))
            {
                int row = Convert.ToInt32(cellrange.Split('-')[0]);
                int col = Convert.ToInt32(cellrange.Split('-')[1]);

                this.range = (Excel.Range)this.worksheet.get_Range(this.GetRowAndColValue(row, col), this.GetRowAndColValue(row, col));
            }
        }

        public void ChangeColumnWidth(string colrange, double colwidth)
        {
            if (this.worksheet == null) return;
            this.range = (Excel.Range)this.worksheet.Columns[colrange];
            this.range.ColumnWidth = colwidth;
        }

        public void ChangeRowHeight(string rowrange, double rowheight)
        {
            if (this.worksheet == null) return;
            this.range = (Excel.Range)this.worksheet.Rows[rowrange];
            this.range.RowHeight = rowheight;
        }

        public void ChangeFont(string cellrange, int size, string name, bool bold)
        {
            this.SelectRange(cellrange);
            if (this.range == null) return;
            if (size != 0) this.range.Font.Size = size;
            if (name != "") this.range.Font.Name = name;
            this.range.Font.Bold = bold;
        }

        public void ChangeFont(string cellrange, int size, string name)
        {
            this.range = (Excel.Range)worksheet.get_Range(cellrange.Split(',')[0], cellrange.Split(',')[1]);
            if (this.range == null) return;
            if (size != 0) this.range.Font.Size = size;
            if (name != "") this.range.Font.Name = name;
        }

        public void ChangeMergeCells(string cellrange, bool merge)
        {
            this.SelectRange(cellrange);
            if (this.range == null) return;
            this.range.MergeCells = merge;
        }

        public void ChangeAlignment(string cellrange, object hori, object vert)
        {
            this.SelectRange(cellrange);
            if (this.range == null) return;
            if (hori != null) this.range.HorizontalAlignment = (Excel.XlHAlign)hori;
            if (vert != null) this.range.VerticalAlignment = (Excel.XlVAlign)vert;
        }

        public void WriteValue(long row, long col, string value, char separe = ';')
        {
            int idx = 0;
            try
            {
                object[] vector = value.Split(separe);
                idx = vector.Length - 1;
                var startCell = (Excel.Range)this.worksheet.Cells[row, col];
                var endCell = (Excel.Range)this.worksheet.Cells[row, col + idx];
                var writeRange = this.worksheet.get_Range(startCell, endCell);
                writeRange.set_Value(Type.Missing, vector);
                vector = null;
                startCell = null;
                endCell = null;
                writeRange = null;
            }
            catch
            {
                idx = 0;
                foreach (string cellvalue in value.Split(';'))
                {
                    try
                    {
                        this.worksheet.Cells[row, col + idx] = cellvalue;
                    }
                    catch
                    {
                    }
                    idx++;
                }
            }
        }

        public void WriteShtValue(long row, long col, string value, string sheetname)
        {
            Excel.Worksheet tmpsht = (Excel.Worksheet)workbook.Sheets[sheetname.ToString()];
            int idx = 0;
            foreach (string cellvalue in value.Split(';'))
            {
                tmpsht.Cells[row, col + idx] = cellvalue;
                idx++;
            }
        }

        public void PaintCell(int row, int col, string refcolor)
        {
            this.SelectRange(row.ToString() + "-" + col.ToString());
            if (refcolor == "Red")
                this.range.Interior.Color = Excel.XlRgbColor.rgbRed;
            else if (refcolor == "Green")
                this.range.Interior.Color = Excel.XlRgbColor.rgbGreen;
            else
                this.range.Interior.Color = 0;
        }

        public void PaintCell(string cellrange, string refcolor)
        {
            this.SelectRange(cellrange);
            if (refcolor == "Red")
                this.range.Interior.Color = Excel.XlRgbColor.rgbRed;
            else if (refcolor == "Green")
                this.range.Interior.Color = Excel.XlRgbColor.rgbGreen;
            else
                this.range.Interior.Color = 0;
        }

        public string GetFileNameFromDialog(bool seleccionmultiple = true)
        {
            string result = "";
            System.Windows.Forms.OpenFileDialog file = new System.Windows.Forms.OpenFileDialog
            {
                Title = "Seleccionar planilla de excel",
                Filter = "Excel (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm|All Files (*.*)|*.*",
            };
            file.Multiselect = seleccionmultiple;

            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string sel in file.FileNames)
                {
                    if (result == "")
                        result = sel;
                    else
                        result = result + ";" + sel;
                }
                file = null;
                return result;
            }
            file = null;
            return result;
        }

        public void pasteRange2(string init_col, string init_row, int cant_col, int cant_row, object[] a_data)
        {
            app.Range[init_col + init_row].Resize[cant_col, cant_row].Value = a_data;
        }

        public void SelectSheet2(string shtname)
        {
            MessageBox.Show("Sheet name is: " + shtname);

            try
            {
                this.worksheet = (Excel.Worksheet)app.Worksheets[shtname.ToString()];
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception is: " + e.ToString());
            }
            
        }

        public string getCell(string sheet, string row, string col)
        {
            worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

            string s_data = string.Empty;
            s_data = worksheet.get_Range("A1", "A1").Value.ToString();
            return s_data;
        }

        public void PrintHeader()
        {
            this.worksheet.PageSetup.PrintArea = "A1:N82";
            this.worksheet.PageSetup.Zoom = 63;
            this.worksheet.PageSetup.HeaderMargin = workbook.Application.CentimetersToPoints(0.8);
            this.worksheet.PageSetup.LeftMargin = workbook.Application.CentimetersToPoints(1.8);
            this.worksheet.PageSetup.RightMargin = workbook.Application.CentimetersToPoints(1.8);
            this.worksheet.PageSetup.TopMargin = workbook.Application.CentimetersToPoints(1.9);
            this.worksheet.PageSetup.BottomMargin = workbook.Application.CentimetersToPoints(1.9);
            this.worksheet.PageSetup.FooterMargin = workbook.Application.CentimetersToPoints(0.8);
            this.worksheet.PageSetup.ScaleWithDocHeaderFooter = true;
            this.worksheet.PageSetup.AlignMarginsHeaderFooter = true;
            this.worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperLetter;
            this.worksheet.Application.ActiveWindow.View = Excel.XlWindowView.xlPageBreakPreview;
            this.worksheet.PageSetup.PrintArea = "A1:N82";
            this.worksheet.PageSetup.Zoom = 63;
            this.worksheet.Application.ActiveWindow.Zoom = 100;


            this.range = (Excel.Range)worksheet.Columns["A:M"];
            this.range.NumberFormat = "@";
            this.range.Font.Name = "Arial";
            this.range.Font.Size = 7;
            this.range.WrapText = false;

            this.ChangeRowHeight("1:2", 14.40);
            this.ChangeRowHeight("3:5", 16.20);
            this.ChangeRowHeight("6:7", 18.60);
            this.ChangeRowHeight("8:77", 12.60);
            this.ChangeRowHeight("78:78", 16.20);
            this.ChangeRowHeight("79:79", 15);
            this.ChangeRowHeight("80:81", 21);


            this.ChangeAlignment("B4,B4", Excel.XlHAlign.xlHAlignRight, null);
            this.ChangeAlignment("B6,L77", Excel.XlHAlign.xlHAlignCenter, Excel.XlHAlign.xlHAlignCenter);
            this.ChangeAlignment("M6,M7", Excel.XlHAlign.xlHAlignCenter, Excel.XlHAlign.xlHAlignCenter);


            this.ChangeColumnWidth("A:A", 6.33);
            this.ChangeFont("A1,A100", 12, "", true);

            this.ChangeAlignment("A1,A100", Excel.XlHAlign.xlHAlignRight, Excel.XlHAlign.xlHAlignCenter);

            this.ChangeColumnWidth("B:B", 10.67);
            this.ChangeColumnWidth("C:C", 8.56);
            this.ChangeColumnWidth("D:D", 9.89);
            this.ChangeColumnWidth("E:E", 7.78);
            this.ChangeColumnWidth("F:F", 8.56);
            this.ChangeColumnWidth("G:G", 9.89);
            this.ChangeColumnWidth("H:H", 7.78);
            this.ChangeColumnWidth("I:I", 12.11);
            this.ChangeColumnWidth("J:J", 7.78);
            this.ChangeColumnWidth("K:K", 7.11);
            this.ChangeColumnWidth("L:L", 7.11);
            this.ChangeColumnWidth("M:M", 10.67);
            this.ChangeColumnWidth("N:N", 4.89);

            this.ChangeFont("B3,B4", 12, "", true);
            this.ChangeFont("C3,C3", 12, "");

            this.ChangeFont("C4,C4", 12, "Arial Narrow", false);

            this.ChangeAlignment("C4,C4", Excel.XlHAlign.xlHAlignLeft, null);

            this.ChangeFont("B6,M7", 10, "");
            this.ChangeFont("B78,B78", 11, "");
            this.ChangeFont("C78,M78", 10, "", true);

            this.ChangeAlignment("C78,M78", Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignCenter);

            this.ChangeFont("E80,H80", 12, "", true);

            this.ChangeFont("I79,J79", 6, "");

            this.ChangeFont("K79,L81", 8, "");
            this.ChangeAlignment("K79,L81", Excel.XlHAlign.xlHAlignCenter, Excel.XlVAlign.xlVAlignCenter);

            this.ChangeFont("M79,M81", 9, "", true);
            this.ChangeAlignment("M79,M81", Excel.XlHAlign.xlHAlignLeft, Excel.XlVAlign.xlVAlignCenter);
            this.range.NumberFormat = "General";

            this.ChangeFont("I80,J81", 20, "", true);


            this.ChangeMergeCells("B6,B7", true);
            this.ChangeMergeCells("C6,E6", true);
            this.ChangeMergeCells("F6,H6", true);
            this.ChangeMergeCells("K6,L6", true);
            this.ChangeMergeCells("M6,M7", true);
            this.ChangeMergeCells("C78,M78", true);
            this.ChangeMergeCells("B79,D81", true);
            this.ChangeMergeCells("E80,H80", true);
            this.ChangeMergeCells("I79,J79", true);
            this.ChangeMergeCells("I80,J81", true);
            this.ChangeMergeCells("K79,L79", true);
            this.ChangeMergeCells("K80,L80", true);
            this.ChangeMergeCells("K81,L81", true);

            this.WriteValue(3, 2, "Proyecto");
            this.WriteValue(4, 2, "Título  ");

            this.WriteValue(6, 2, "Señal");
            this.WriteValue(6, 3, "Desde");
            this.WriteValue(6, 6, "Hasta");
            this.WriteValue(6, 9, "Sección");
            this.WriteValue(6, 10, "Longitud");
            this.WriteValue(6, 11, "Hojas Funcional");
            this.WriteValue(6, 13, "Chequeo");

            this.WriteValue(7, 3, "Ubicación");
            this.WriteValue(7, 4, "Bornera");
            this.WriteValue(7, 5, "Terminal");
            this.WriteValue(7, 6, "Ubicación");
            this.WriteValue(7, 7, "Bornera");
            this.WriteValue(7, 8, "Terminal");
            this.WriteValue(7, 10, "[ m ]");
            this.WriteValue(7, 11, "Desde");
            this.WriteValue(7, 12, "Hasta");

            this.WriteValue(78, 2, "TITULO:");
            this.WriteValue(78, 3, "PLANILLA DE MODIFICACIONES");
            this.WriteValue(79, 9, "REVISION");
            this.WriteValue(79, 11, "HOJA N°:");
            this.WriteValue(80, 11, "CONT. EN:");
            this.WriteValue(81, 11, "FECHA:");

            //Logo
            this.WriteValue(79, 2, "ABB");
            this.range = worksheet.get_Range("B79", "D81");
            this.range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Font.Size = 36;
            this.range.Font.Name = "ABB Logo Bold";
            this.range.Font.ColorIndex = 3;


            this.range = worksheet.get_Range("B6", "M77");
            this.range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders.Weight = Excel.XlBorderWeight.xlThin;

            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThick;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThick;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

            this.range = worksheet.get_Range("B6", "M7");
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;

            this.range = worksheet.get_Range("C6", "E77");
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;

            this.range = worksheet.get_Range("F6", "H77");
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;

            this.range = worksheet.get_Range("I6", "I77");
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;

            this.range = worksheet.get_Range("K6", "L77");
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;

            this.range = worksheet.get_Range("B79", "D81");
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            this.range.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

            this.range = worksheet.get_Range("I6", "J7");
            this.range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;


            this.range = worksheet.get_Range("E79", "H81");
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            this.range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            this.range = worksheet.get_Range("I79", "J81");
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            this.range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            this.range = worksheet.get_Range("K79", "M81");
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlMedium;

        }

        private void SetFormatting(Excel.Range range)
        {
            this.range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
            this.range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            this.range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            this.range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
        }

        public void CompareExcels(string strold, string strnew)
        {
            ExcelInterface xlsapp_old = new ExcelInterface();
            ExcelInterface xlsapp_new = new ExcelInterface();
            object[,] xlsrange_old = null;
            object[,] xlsrange_new = null;
            Excel.Worksheet worksheet;
            Excel.Range last;
            Excel.Range range;
            string celldir = string.Empty;

            try
            {
                foreach (Excel.Worksheet xlShtOld in xlsapp_new.workbook.Worksheets)
                {
                    foreach (Excel.Worksheet xlShtNew in xlsapp_old.workbook.Worksheets)
                    {
                        if (xlShtNew.Name == xlShtOld.Name)
                        {
                            last = xlShtNew.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                            range = xlShtNew.get_Range("A1", last);
                            xlsrange_new = range.Text as object[,];

                            last = xlShtOld.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                            range = xlShtOld.get_Range("A1", last);
                            xlsrange_old = range.Text as object[,];

                            Parallel.For(1, xlsrange_new.GetLength(0) + 1, row =>
                            {
                                Parallel.For(1, xlsrange_new.GetLength(1) + 1, col =>
                                {
                                    if (this.GetData(xlsrange_new[row, col]) != this.GetData(xlsrange_old[row, col]))
                                    {
                                        worksheet = (Excel.Worksheet)xlsapp_new.workbook.Sheets[xlShtNew.Name];
                                        celldir = GetRowAndColValue(row, col);
                                        worksheet.get_Range(celldir, celldir).Interior.Color = Excel.XlRgbColor.rgbRed;
                                    }
                                });
                            });

                        }
                    }
                }
            }
            catch
            {
            }
            finally
            {
                worksheet = null;
                last = null;
                range = null;
                xlsrange_old = null;
                xlsrange_new = null;
                xlsapp_old.CloseXLS();
                xlsapp_old.ReleaseWithoutCloseXLS();
                xlsapp_old = null;
                xlsapp_new.ReleaseWithoutCloseXLS();
                xlsapp_new = null;
                GC.Collect();
            }
        }

        public List<List<String>> ExcelRangeToLists(Excel.Range cells)
        {
            return cells.Rows.Cast<Excel.Range>().AsParallel().Select(row =>
            {
                return row.Cells.Cast<Excel.Range>().Select(cell =>
                {
                    var cellContent = cell.Value2;
                    return (cellContent == null) ? string.Empty : cellContent.ToString();
                }).Cast<string>().ToList();
            }).ToList();
        }


        public bool WriteNumber(long row, long col, int value)
        {
            try
            {
                this.worksheet.Cells[row, col] = value;
            }
            catch
            {
                return false;
            }
            return true;
        }
        
        //public string GetColumnName (int idx)
        //{
        //    const int AlphabetLength = 'Z' - 'A' + 1;
        //    string columnName = l_colend1.ToUpperInvariant();

        //    if (idx != 1)
        //    {
        //        int columnNumber = 0;
        //        for (int i = 0; i < columnName.Length; i++)
        //        {
        //            columnNumber *= AlphabetLength;
        //            columnNumber += (columnName[i] - 'A' + 1);
        //        }

        //        columnName = "";
        //        columnNumber = columnNumber + idx - 1;

        //        int dividend = columnNumber;
        //        int modulo;
        //        while (dividend > 0)
        //        {
        //            modulo = (dividend - 1) % AlphabetLength;
        //            columnName = Convert.ToChar('A' + modulo).ToString() + columnName;
        //            dividend = (int)((dividend - modulo) / AlphabetLength);
        //        }

        //    }
        //    return columnName;
        //}

    }
}
