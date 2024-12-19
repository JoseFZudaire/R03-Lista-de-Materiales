//Generate cs correct file


using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Interface.ExcelTools;
using ListOfMaterialsPGGA.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ListOfMaterialsPGGA
{
    public partial class Generate : Form
    {
        //const Int32 blank_color = 0xFFFFFF;
        const Int32 blank_color = 0xD58D53;
        const Int32 warning_color = 0x00FFFF;
        const Int32 new_row_color = 0xE0A413; // #13a4e0
        const string quote = "\"";
        const int col_item = 1;
        const int col_desc = 2;
        const int col_codefab = 7;
        const int col_need = 12;
        const int col_need_total = 8;
        const int col_build = 12;
        const int col_saldo = 14;
        const int col_resp = 15;

        //Excel.Application app = new Excel.Application();

        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

        public Generate()
        {
            InitializeComponent();
        }

        private void Generate_Load(object sender, EventArgs e)
        {
            //using some code here to step on the template file, because I really dislike using the Excel Interface code

            string oldPathName = AppDomain.CurrentDomain.BaseDirectory + "Template_2.xlsm";
            string newPathName = AppDomain.CurrentDomain.BaseDirectory + "Template.xlsm";

            try
            {
                System.IO.File.Copy(oldPathName, newPathName, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            this.picbox_save.Image = new Bitmap(Properties.Resources.uncheck, new Size(24, 24));
            this.tb_template.Text = AppDomain.CurrentDomain.BaseDirectory + "Template.xlsm";
            this.tb_path.Text = AppDomain.CurrentDomain.BaseDirectory;
            this.tb_save.Text = "Report-" + DateTime.Now.ToString("yyyy-MM-dd") + "-BillOfMaterial.xlsm";
            this.st_label.Text = "Inicio del programa";
            this.mensaje.Update();
        }

        private void Fill_template()
        {
            this.st_label.Text = "";
            this.mensaje.Update();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm|All Files (*.*)|*.*";
            ofd.FilterIndex = 2;

            DialogResult result = ofd.ShowDialog();

            if (result == DialogResult.OK)
            {
                this.tb_template.Text = ofd.FileName;
            }
            ofd = null;
        }

        private void Fill_save()
        {
            this.st_label.Text = "";
            this.mensaje.Update();
            FolderBrowserDialog ofd = new FolderBrowserDialog();

            DialogResult result = ofd.ShowDialog();

            if (result == DialogResult.OK)
            {
                this.tb_path.Text = ofd.SelectedPath + "\\";
            }
            ofd = null;
        }

        private void bt_add_Click(object sender, EventArgs e)
        {
            this.st_label.Text = "";
            this.mensaje.Update();
            ExcelInterface obj = new ExcelInterface();
            string excelfilepath = string.Empty;
            excelfilepath = obj.GetFileNameFromDialog();

            List<string> shts_not_to_include = new List<string>() { "Cover page", "BoMListMechanical", "Caratula", "Total", "Totales", "ComponentData",
                                                                "Referencia", "Lista de Repuestos", "Lista de Sueltos", "Panel de Control", "Modelo"};

            obj.Dispose();
            obj = null;

            bool firstRowAdded = false;

            app.Visible = false;
            app.ScreenUpdating = false;
            app.DisplayAlerts = false;

            foreach (string selectedfile in excelfilepath.Split(';'))
            {
                if (selectedfile.Split('\\').Last() == tb_save.Text)
                {
                    MessageBox.Show("No se puede agregar un documento con el mismo nombre con el que va a ser guardado");
                }

                //IN REALITY YOU HAVE OPEN DOCUMENT AND CHECK IF ITS JUST BOMLIST OR IF IT HAS MULTIPLE PAGES, AND CHECK FOR 
                //EACH SPECIFICALLY, AND ADD THEM

                else if(selectedfile != "")
                {
                    Workbook book = app.Workbooks.Open(selectedfile,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

                    //Excel.Workbook book = app.Workbooks.Open(selectedfile);

                    foreach(Excel.Worksheet ws in book.Sheets)
                    {

                        if (dataGridView1.RowCount < 1)
                        {
                            dataGridView1.Rows.Add();
                            firstRowAdded = true;
                        }

                        List<string> files_added = new List<string>();

                        foreach (var item in this.dataGridView1.Rows)
                        {
                            files_added.Add((string)((DataGridViewRow)item).Cells[0].Value);
                        }

                        DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                        row.Cells[0].Value = ws.Name;
                        row.Cells[1].Value = System.IO.Path.GetFileName(selectedfile);
                        row.Cells[3].Value = selectedfile;

                        if ((!(files_added.Contains(ws.Name)) && !(shts_not_to_include.Contains(ws.Name))) || (ws.Name == "BoMList"))
                        {
                            this.dataGridView1.Rows.Add(row);
                            this.st_label.Text = "Documento agregado a la lista";
                            this.mensaje.Update();
                        }
                        else if(files_added.Contains(ws.Name))
                        {
                            MessageBox.Show("No se puede incluir el archivo " + ws.Name + " porque ya existe ese nombre.", "Agregar archivo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

                    List<string> addedFiles = new List<string>();

                    foreach (var item in this.dataGridView1.Rows)
                    {
                        addedFiles.Add((string)((DataGridViewRow)item).Cells[0].Value);
                    }

                    //book.Save();

                    System.Threading.Thread.Sleep(100);
                    book.Close();
                }
            }

            System.Threading.Thread.Sleep(100);
            //app.Quit();

            if (firstRowAdded)
            {
                dataGridView1.Rows.RemoveAt(0);
            }

            List<string> nameFiles = new List<string>();

            nameFiles.Add("");

            foreach (var item in this.dataGridView1.Rows)
            {
                if ((string)((DataGridViewRow)item).Cells[1].Value != "Hoja vacía")
                {
                    if ((string)((DataGridViewRow)item).Cells[0].Value == "BoMList")
                    {
                        nameFiles.Add((string)((DataGridViewRow)item).Cells[1].Value);
                    }
                    else
                    {
                        nameFiles.Add((string)((DataGridViewRow)item).Cells[0].Value);
                    }
                }
            }

            for (int j = 0; j < this.dataGridView1.Rows.Count; j++)
            {
                string combineWith = (string)(((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2].Value);
                ((DataGridViewComboBoxCell)((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2]).Items.Clear();

                foreach (var option in nameFiles)
                {
                    if ((option != (((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[0].Value).ToString()) &&
                        ((option != (((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[1].Value).ToString()) ||
                        (((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[0].Value).ToString() != "BoMList"))
                    {
                        ((DataGridViewComboBoxCell)((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2]).Items.Add(option);
                    }
                }

                if (nameFiles.Contains(combineWith))
                {
                    ((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2].Value = combineWith;
                }
                else
                {
                    ((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2].Value = "";
                }
            }

            app.DisplayAlerts = true;

        }

        private void Generate_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }


        private void bt_del_Click(object sender, EventArgs e)
        {
            this.st_label.Text = "";
            this.mensaje.Update();
            DataGridView dgv = this.dataGridView1;
            try
            {
                int totalRows = dgv.Rows.Count;
                if (totalRows == 0)
                    return;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                dgv.Rows.Remove(selectedRow);
                dgv.ClearSelection();
                this.st_label.Text = "Documento borrado de la lista";
                this.mensaje.Update();
                if (dgv.Rows.Count == 0)
                    return;
                dgv.Rows[0].Selected = true;
            }
            catch { }
            dgv = null;

            List<string> nameFiles = new List<string>();

            nameFiles.Add("");

            foreach (var item in this.dataGridView1.Rows)
            {
                if ((string)((DataGridViewRow)item).Cells[1].Value != "Hoja vacía")
                {
                    if ((string)((DataGridViewRow)item).Cells[0].Value == "BoMList")
                    {
                        nameFiles.Add((string)((DataGridViewRow)item).Cells[1].Value);
                    }
                    else
                    {
                        nameFiles.Add((string)((DataGridViewRow)item).Cells[0].Value);
                    }
                }
            }

            for (int j = 0; j < this.dataGridView1.Rows.Count; j++)
            {
                string combineWith = (string)(((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2].Value);
                ((DataGridViewComboBoxCell)((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2]).Items.Clear();

                foreach (var option in nameFiles)
                {
                    if ((option != (((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[0].Value).ToString()) &&
                        ((option != (((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[1].Value).ToString()) ||
                        (((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[0].Value).ToString() != "BoMList"))
                    {
                        ((DataGridViewComboBoxCell)((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2]).Items.Add(option);
                    }
                }

                if (nameFiles.Contains(combineWith))
                {
                    ((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2].Value = combineWith;
                }
                else
                {
                    ((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2].Value = "";
                }
            }


            for (int k = 0; k < dataGridView1.Rows.Count; k++)
            {
                if (dataGridView1.Rows[k].Cells[2].ReadOnly)
                {
                    bool match = false;
                    int col = 0;

                    if (dataGridView1.Rows[k].Cells[0].Value.ToString() == "BoMList")
                    {
                        col = 1;
                    }

                    for (int n = 0; n < dataGridView1.Rows.Count; n++)
                    {
                        if ((k != n) && (dataGridView1.Rows[n].Cells[2].Value != null) && (dataGridView1.Rows[n].Cells[2].Value != DBNull.Value) &&
                            !(String.IsNullOrWhiteSpace((dataGridView1.Rows[n].Cells[2].Value).ToString())))
                        {
                            if (dataGridView1.Rows[n].Cells[2].Value.ToString() == dataGridView1.Rows[k].Cells[col].Value.ToString())
                            {
                                match = true;
                            }
                        }
                    }

                    if (dataGridView1.Rows[k].Cells[1].Value.ToString() == "Hoja vacía")
                    {
                        match = true;
                    }

                    if (!match)
                    {
                        dataGridView1.Rows[k].Cells[2].ReadOnly = false;
                        dataGridView1.Rows[k].DefaultCellStyle.BackColor = Color.White;
                        dataGridView1.Refresh();
                    }
                }
            }
        }

        private void bt_up_Click(object sender, EventArgs e)
        {
            this.st_label.Text = "";
            this.mensaje.Update();
            DataGridView dgv = this.dataGridView1;
            try
            {
                int totalRows = dgv.Rows.Count;
                if (totalRows < 2)
                    return;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                if (rowIndex == 0)
                    return;
                // get index of the column for the selected cell
                int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                dgv.Rows.Remove(selectedRow);
                dgv.Rows.Insert(rowIndex - 1, selectedRow);
                dgv.ClearSelection();
                dgv.Rows[rowIndex - 1].Cells[colIndex].Selected = true;
            }
            catch { }
            dgv = null;
        }

        private void bt_down_Click(object sender, EventArgs e)
        {
            this.st_label.Text = "";
            this.mensaje.Update();
            DataGridView dgv = this.dataGridView1;
            try
            {
                int totalRows = dgv.Rows.Count;
                if (totalRows < 2)
                    return;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                if (rowIndex == totalRows - 1)
                    return;
                // get index of the column for the selected cell
                int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                dgv.Rows.Remove(selectedRow);
                dgv.Rows.Insert(rowIndex + 1, selectedRow);
                dgv.ClearSelection();
                dgv.Rows[rowIndex + 1].Cells[colIndex].Selected = true;
            }
            catch { }
            dgv = null;
        }

        private void bt_Sumar_Click(object sender, EventArgs e)
        {
            this.st_label.Text = "";
            this.mensaje.Update();
            DataGridView dgv = this.dataGridView1;
            try
            {
                int totalRows = dgv.Rows.Count;
                if (totalRows == 0)
                    return;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                int value = Convert.ToInt32(selectedRow.Cells[2].Value);
                value++;
                selectedRow.Cells[2].Value = value;
            }
            catch { }
            dgv = null;
        }

        private void bt_Restar_Click(object sender, EventArgs e)
        {
            this.st_label.Text = "";
            this.mensaje.Update();
            DataGridView dgv = this.dataGridView1;
            try
            {
                int totalRows = dgv.Rows.Count;
                if (totalRows == 0)
                    return;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                int value = Convert.ToInt32(selectedRow.Cells[2].Value);
                value--;
                if (value > 0)
                    selectedRow.Cells[2].Value = value;
            }
            catch { }
            dgv = null;
        }

        private void bt_template_Click(object sender, EventArgs e)
        {
            Fill_template();
        }

        private void bt_save_Click(object sender, EventArgs e)
        {
            Fill_save();
        }

        private void tb_template_DoubleClick(object sender, EventArgs e)
        {
            Fill_template();
        }

        private void tb_path_DoubleClick(object sender, EventArgs e)
        {
            Fill_save();
        }

        private void tb_template_TextChanged(object sender, EventArgs e)
        {

        }

        private void tb_save_TextChanged(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(System.IO.Path.Combine(this.tb_path.Text, this.tb_save.Text)))
                this.picbox_save.Image = new Bitmap(Properties.Resources.check, new Size(24, 24));
            else
                this.picbox_save.Image = new Bitmap(Properties.Resources.uncheck, new Size(24, 24));
        }

        private void bt_out_open_Click(object sender, EventArgs e)
        {
            try
            {
                string filepath = this.tb_path.Text;
                Boolean found = false;
                System.Diagnostics.ProcessStartInfo StartInformation = new System.Diagnostics.ProcessStartInfo();
                while (!found)
                {
                    StartInformation.FileName = filepath;
                    try
                    {
                        System.Diagnostics.Process process = System.Diagnostics.Process.Start(StartInformation);
                        process = null;
                        found = true;
                    }
                    catch { }
                    int ultimo = filepath.LastIndexOf('\\');
                    filepath = filepath.Remove(ultimo);
                }
                StartInformation = null;
            }
            catch { }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            for(int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if ((i != e.RowIndex)&&(dataGridView1.Rows[e.RowIndex].Cells[2].Value != null)&&(dataGridView1.Rows[e.RowIndex].Cells[2].Value != DBNull.Value)&&
                    !(String.IsNullOrWhiteSpace((dataGridView1.Rows[e.RowIndex].Cells[2].Value).ToString())))
                {
                    int col = 0;

                    if (dataGridView1.Rows[i].Cells[0].Value.ToString() == "BoMList")
                    {
                        col = 1;
                    }

                    if (dataGridView1.Rows[i].Cells[col].Value.ToString() == dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString())
                    {
                        dataGridView1.Rows[i].Cells[2].Value = "";
                        dataGridView1.Rows[i].Cells[2].ReadOnly = true;
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGray;
                        dataGridView1.Refresh();
                        break;
                    }
                }
            }

            for (int k = 0; k < dataGridView1.Rows.Count; k++)
            {
                if (dataGridView1.Rows[k].Cells[2].ReadOnly)
                {
                    bool match = false;
                    int col = 0;

                    if (dataGridView1.Rows[k].Cells[0].Value.ToString() == "BoMList")
                    {
                        col = 1;
                    }

                    for (int n = 0; n < dataGridView1.Rows.Count; n++)
                    {
                        if((k!=n) && (dataGridView1.Rows[n].Cells[2].Value != null) && (dataGridView1.Rows[n].Cells[2].Value != DBNull.Value) &&
                            !(String.IsNullOrWhiteSpace((dataGridView1.Rows[n].Cells[2].Value).ToString())))
                        {
                            if(dataGridView1.Rows[n].Cells[2].Value.ToString() == dataGridView1.Rows[k].Cells[col].Value.ToString())
                            {
                                match = true;
                            }
                        }
                    }

                    if (dataGridView1.Rows[k].Cells[1].Value.ToString() == "Hoja vacía")
                    {
                        match = true;
                    }

                    if (!match)
                    {
                        dataGridView1.Rows[k].Cells[2].ReadOnly = false;
                        dataGridView1.Rows[k].DefaultCellStyle.BackColor = Color.White;
                        dataGridView1.Refresh();
                    }
                }
            }
        }


        private void bt_generate_Click(object sender, EventArgs e)
        {
            string oldPathName = AppDomain.CurrentDomain.BaseDirectory + "Template_2.xlsm";
            string newPathName = AppDomain.CurrentDomain.BaseDirectory + "Template.xlsm";

            try
            {
                System.IO.File.Copy(oldPathName, newPathName, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //abrir el template, grabar en un nuevo archivo
            this.st_label.Text = "Proceso para generar nuevo documento...";
            this.mensaje.Update();
            if (this.dataGridView1.Rows.Count < 1)
            {
                this.st_label.Text = "La lista no contiene documentos.";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!System.IO.File.Exists(this.tb_template.Text))
            {
                this.st_label.Text = "No se encontró el archivo template.";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
                return;
            }
            if (this.tb_template.Text == (System.IO.Path.Combine(this.tb_path.Text, this.tb_save.Text)))
            {
                this.st_label.Text = "No debe usar el archivo de template como salida.";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            this.st_label.Text = "Abriendo archivo template...";
            this.mensaje.Update();
            ExcelInterface objxls = null;
            try
            {
                objxls = new ExcelInterface();
                objxls.OpenFileToEdit(this.tb_template.Text);
                objxls.app.DisplayAlerts = false;
            }
            catch
            {
                this.st_label.Text = "Error al conectarse con el programa Excel";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            //if (!objxls.SelectSheet("Referencia"))
            //{
            //    if (!objxls.SelectSheet("Total"))
            //    {
            //        objxls.CloseXLS();
            //        objxls.ReleaseWithoutCloseXLS();
            //        objxls = null;
            //        this.st_label.Text = "El archivo template no contiene una hoja con nombre 'Total'.";
            //        this.mensaje.Update();
            //        MessageBox.Show(this.st_label.Text);
            //        return;
            //    }
            //}
            objxls.HideExcel();

            objxls.app.DisplayAlerts = false;
            objxls.TrySaveAs(System.IO.Path.Combine(this.tb_path.Text, this.tb_save.Text));
            //MessageBox.Show("Save path: " + this.tb_path.Text + this.tb_save.Text);
            if (!System.IO.File.Exists(System.IO.Path.Combine(this.tb_path.Text, this.tb_save.Text)))
            {
                this.st_label.Text = "Problema al intentar guardar el archivo.";
                this.mensaje.Update();
                this.picbox_save.Image = new Bitmap(Properties.Resources.uncheck, new Size(24, 24));
                objxls.CloseXLS();
                objxls.ReleaseWithoutCloseXLS();
                objxls = null;
                MessageBox.Show(this.st_label.Text, "Problema al guardar archivo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            objxls.HideExcel();
            this.st_label.Text = "Archivo base creado correctamente, generando hojas...";
            this.mensaje.Update();
            this.picbox_save.Image = new Bitmap(Properties.Resources.check, new Size(24, 24));

            string filedir = string.Empty;
            string tempname = string.Empty;
            //recorrer los archivos de la lista, ir copiando toda la info en una nueva hoja
            //cada archivo de la lista debe ser abierto en modo lectura, sacar la info y cerrarlo
            //llenar una lista con cada archivo que se fue copiando
            int count = 1;

            try
            {
                for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
                {
                    if(!(this.dataGridView1.Rows[i].Cells[2].ReadOnly))
                    {
                        filedir = (string)this.dataGridView1.Rows[i].Cells[3].Value;
                        try
                        {
                            objxls.bookmem = objxls.app.Workbooks.Open(filedir, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                            objxls.app.DisplayAlerts = false;
                            //objxls.app.Visible = true;
                        }
                        catch(Exception exc)
                        {
                            MessageBox.Show(exc.ToString());
                        }

                        if (objxls.bookmem != null)
                        {
                            tempname = (string)this.dataGridView1.Rows[i].Cells[0].Value;

                            foreach (Excel.Worksheet wksht in objxls.bookmem.Worksheets)
                            {
                                if (wksht.Name == tempname)
                                {
                                    //MessageBox.Show("Wksht name : " + wksht.Name);
                                    objxls.sheetmem = wksht;
                                }
                            }

                            if ((this.dataGridView1.Rows[i].Cells[2].Value != null) && (this.dataGridView1.Rows[i].Cells[2].Value != DBNull.Value)
                                && !(String.IsNullOrWhiteSpace(this.dataGridView1.Rows[i].Cells[2].Value.ToString())) && (this.dataGridView1.Rows[i].Cells[2].Value.ToString() != ""))
                            {
                                this.st_label.Text = "Copiando " + count + "/" + this.dataGridView1.RowCount + " ...";
                                this.mensaje.Update();

                                objxls.SelectSheet("Referencia");

                                objxls.sheetmem.Copy(Type.Missing, (objxls.worksheet));

                                string originalSht = (string)this.dataGridView1.Rows[i].Cells[0].Value;

                                if (tempname == "BoMList")
                                {
                                    originalSht = (string)this.dataGridView1.Rows[i].Cells[1].Value;
                                } 

                                objxls.SelectSheet(tempname);
                                objxls.worksheet.Name = "Duplicate";

                                for (int j = 0; j < this.dataGridView1.Rows.Count; j++)
                                {
                                    int col = 0;

                                    if (this.dataGridView1.Rows[j].Cells[0].Value.ToString() == "BoMList")
                                    {
                                        col = 1;
                                    }

                                    string shtName = "";

                                    if(this.dataGridView1.Rows[j].Cells[col].Value.ToString().Length > 28)
                                    {
                                        shtName = this.dataGridView1.Rows[j].Cells[col].Value.ToString().Substring(0, 28);
                                    } else
                                    {
                                        shtName = this.dataGridView1.Rows[j].Cells[col].Value.ToString();
                                    }

                                    if (this.dataGridView1.Rows[j].Cells[col].Value.ToString() == this.dataGridView1.Rows[i].Cells[2].Value.ToString())
                                    {
                                        string filedir_combine = (string)this.dataGridView1.Rows[j].Cells[3].Value;
                                        objxls.bookmem.Close(false, false, true);
                                        objxls.bookmem = objxls.app.Workbooks.Open(filedir_combine, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                                        objxls.app.DisplayAlerts = false;

                                        foreach (Excel.Worksheet wksht_combine in objxls.bookmem.Worksheets)
                                        {
                                            if (wksht_combine.Name == this.dataGridView1.Rows[j].Cells[0].Value.ToString())
                                            {
                                                objxls.sheetmem = wksht_combine;
                                            }
                                        }

                                        objxls.sheetmem.Copy(Type.Missing, (objxls.worksheet));
                                        objxls.SelectSheet(this.dataGridView1.Rows[j].Cells[0].Value.ToString());
                                        objxls.worksheet.Name = shtName;

                                        objxls = groupSheets(objxls, shtName, originalSht);
                                        //groupSheets(objxls, shtName);

                                        break;
                                    }
                                }
                            }
                            else
                            {
                                this.st_label.Text = "Copiando " + count + "/" + this.dataGridView1.RowCount + " ...";
                                this.mensaje.Update();

                                objxls.SelectSheet("Referencia");

                                objxls.sheetmem.Copy(Type.Missing, (objxls.worksheet));

                                if (tempname == "BoMList")
                                {
                                    objxls.SelectSheet(tempname);

                                    string shtName = "";

                                    if (this.dataGridView1.Rows[i].Cells[1].Value.ToString().Length > 28)
                                    {
                                        shtName = this.dataGridView1.Rows[i].Cells[1].Value.ToString().Substring(0, 28);
                                    }
                                    else
                                    {
                                        shtName = this.dataGridView1.Rows[i].Cells[1].Value.ToString();
                                    }

                                    //objxls.worksheet.Name = this.dataGridView1.Rows[i].Cells[1].Value.ToString();
                                    objxls.worksheet.Name = shtName;
                                }

                                objxls.worksheet.get_Range(objxls.worksheet.Cells[3, 1], objxls.worksheet.Cells[3, 16]).AutoFilter(1);

                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 2d;
                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 2d;
                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 2d;
                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 2d;
                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeTop].Color = 0x000000;
                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;

                                ((Range)objxls.worksheet.get_Range("A200:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3d;
                                ((Range)objxls.worksheet.get_Range("A3:P3")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3d;
                                ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 4d;
                                ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 4d;
                                ((Range)objxls.worksheet.get_Range("O200:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 4d;

                                ((Range)objxls.worksheet.get_Range("A200:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
                                ((Range)objxls.worksheet.get_Range("A3:P3")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
                                ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
                                ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
                                ((Range)objxls.worksheet.get_Range("O200:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;

                                ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Interior.Color = 0xE6E6E6;

                                ((Range)objxls.worksheet.get_Range("A4:A200")).Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                ((Range)objxls.worksheet.get_Range("H4:O200")).Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                                ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                            }

                            //objxls.bookmem.Save();

                            //objxls.bookmem.Close(true, true, false);
                            objxls.bookmem.Close(false, false, true);
                        }   
                    }
                    else if(this.dataGridView1.Rows[i].Cells[1].Value.ToString() == "Hoja vacía")
                    {
                        filedir = (string)this.dataGridView1.Rows[i].Cells[3].Value;
                        try
                        {
                            objxls.bookmem = objxls.app.Workbooks.Open(filedir, 0, false, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
                            objxls.app.Visible = false;
                            objxls.app.DisplayAlerts = false;
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.ToString());
                        }

                        foreach (Excel.Worksheet wksht in objxls.bookmem.Worksheets)
                        {
                            if (wksht.Name == "BoMList")
                            {
                                //MessageBox.Show("Wksht name : " + wksht.Name);
                                objxls.sheetmem = wksht;
                            }
                        }

                        this.st_label.Text = "Copiando " + count + "/" + this.dataGridView1.RowCount + " ...";
                        this.mensaje.Update();

                        objxls.SelectSheet("Referencia");

                        objxls.sheetmem.Copy(Type.Missing, (objxls.worksheet));

                        objxls.SelectSheet("BoMList");

                        string shtName = "";

                        if (this.dataGridView1.Rows[i].Cells[0].Value.ToString().Length > 28)
                        {
                            shtName = this.dataGridView1.Rows[i].Cells[0].Value.ToString().Substring(0, 28);
                        }
                        else
                        {
                            shtName = this.dataGridView1.Rows[i].Cells[0].Value.ToString();
                        }

                        //objxls.worksheet.Name = this.dataGridView1.Rows[i].Cells[1].Value.ToString();
                        objxls.worksheet.Name = shtName;

                        objxls.bookmem.Close(false, false, true);
                    }

                    count++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception " + ex);

                this.st_label.Text = "Error al adquirir datos de los documentos";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
                objxls.CloseXLS();
                objxls.ReleaseWithoutCloseXLS();
                objxls.Dispose();
                objxls = null;
                return;
            }

            try
            {
                objxls.workbook.Save();
                objxls.SelectSheet("Total");
                this.st_label.Text = "Plantilla de proyecto generada";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text, "Generación de BOM", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch
            {
                this.st_label.Text = "Error al intentar guardar el archivo";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
            }

            //objxls.bookmem.Close();
            objxls.app.DisplayAlerts = true;

            objxls.ShowExcel();
            objxls.Dispose();
            objxls = null;
            return;
        }

        private ExcelInterface groupSheets(ExcelInterface objxls, string shtCombine, string originalSht)
        {
            Dictionary<string, Tuple<int, Int32, string, int>> dict_comp = new Dictionary<string, Tuple<int, Int32, string, int>>();
            List<string> sheets = null;

            int blank_el = 0;
            int count = 0; 

            sheets = new List<string> { "Duplicate", shtCombine };

            foreach (string sht in sheets)
            {
                objxls.SelectSheet(sht);

                int n_quantity = 11;
                int n_comments = 16;

                for (int i = 4; i < 200; i++)
                {
                    string gral_desc = "";

                    for (int k = 2; k < 8; k++) //esto lo que hace es crear el identificador en base a las primeras 6 columnas
                    {
                        gral_desc += (objxls.worksheet.Cells[i, k] as Excel.Range).Text.ToString();
                    }

                    if (gral_desc != "") {

                        string key_comp = "";

                        for(int n = 2; n <= 8; n++)
                        {
                            key_comp += (objxls.worksheet.Cells[i, n] as Excel.Range).Text.ToString() + "¬";
                        }

                        key_comp += (objxls.worksheet.Cells[i, 9] as Excel.Range).Text.ToString();
                        //string key_comp = (objxls.worksheet.Cells[i, 2] as Excel.Range).Text.ToString() + "¬" + (objxls.worksheet.Cells[i, 9] as Excel.Range).Text.ToString();

                        if (!dict_comp.ContainsKey(key_comp))
                        {
                            string el_desc = "";

                            for (int j = 3; j < 16; j++) //esto lo que hace es crear el identificador en base a las primeras 6 columnas
                            {
                                el_desc += (objxls.worksheet.Cells[i, j] as Excel.Range).Text.ToString() + "¬";
                            }
                            el_desc += (objxls.worksheet.Cells[i, 16] as Excel.Range).Text.ToString();

                            try
                            {
                                if (sht == "Duplicate")
                                {
                                    //If its the first sheet, then its not highlighted
                                    dict_comp.Add(key_comp, new Tuple<int, Int32, string, int>(Int32.Parse((objxls.worksheet.Cells[i, n_quantity] as Excel.Range).Text.ToString()), blank_color, el_desc, count));
                                }
                                else
                                {
                                    dict_comp.Add(key_comp, new Tuple<int, Int32, string, int>(Int32.Parse((objxls.worksheet.Cells[i, n_quantity] as Excel.Range).Text.ToString()), new_row_color, el_desc, count));
                                }
                            }
                            catch (Exception exc) { string error = exc.ToString(); }
                        }
                        else
                        {
                            try
                            {
                                //If the value is different to the one already in the dictionary => have to highlight it
                                if ((dict_comp[key_comp].Item1 != Int32.Parse((objxls.worksheet.Cells[i, n_quantity] as Excel.Range).Text.ToString())) && (dict_comp[key_comp].Item2 != new_row_color))
                                {
                                    //Simply changing the value of the number of items available
                                    dict_comp[key_comp] = new Tuple<int, Int32, string, int>(Int32.Parse((objxls.worksheet.Cells[i, n_quantity] as Excel.Range).Text.ToString()), warning_color, dict_comp[key_comp].Item3, dict_comp[key_comp].Item4);
                                }
                            }
                            catch (Exception exc) { string error = exc.ToString(); }
                        }
                    }

                    count++;
                }
            }

            objxls.SelectSheet("Duplicate");

            int n_row = 4;

            var sortedValues = dict_comp.OrderBy(kvp => kvp.Value.Item4);

            //foreach (KeyValuePair<string, Tuple<int, Int32, string, int>> entry in dict_comp)
            foreach (KeyValuePair<string, Tuple<int, Int32, string, int>> entry in sortedValues)
            {
                //string[] descriptions = entry.Value.Item3.Split(';');
                string[] descriptions = entry.Value.Item3.Split('¬');

                objxls.worksheet.Cells[n_row, 1] = n_row - 3;

                for (int j = 3; j <= 8; j++) //pasting all the values of the descriptions
                {
                    objxls.worksheet.Cells[n_row, j] = descriptions[j - 3];
                }

                for (int j = 12; j <= 15; j++) //pasting all the values of the descriptions
                {
                    objxls.worksheet.Cells[n_row, j] = descriptions[j - 3];

                    if (descriptions[j - 3] == "\u2713")
                    {
                        objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, j], objxls.worksheet.Cells[n_row, j]).Font.Color = 0x138a1f;
                    }
                    else
                    {
                        objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, j], objxls.worksheet.Cells[n_row, j]).Font.Color = 0x0000FF;
                    }
                }

                objxls.worksheet.Cells[n_row, 16] = descriptions[13];

                //if (!(entry.Key).Contains("blank"))
                //{
                //objxls.worksheet.Cells[n_row, 9] = (entry.Key).Split('¬')[1];
                objxls.worksheet.Cells[n_row, 9] = (entry.Key).Split('¬')[7];

                //objxls.worksheet.Cells[n_row, 2] = (entry.Key).Split('¬')[0];
                objxls.worksheet.Cells[n_row, 2] = (entry.Key).Split('¬')[0];
                //} else
                //{
                //    objxls.worksheet.Cells[n_row, 9] = "";
                //}

                //objxls.worksheet.Cells[n_row, 16] = descriptions[descriptions.Length - 1]; //comments

                objxls.worksheet.Cells[n_row, 11] = (entry.Value.Item1).ToString(); //pasting the value of the quantity

                if (entry.Value.Item2 == new_row_color)
                {
                    objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, 1], objxls.worksheet.Cells[n_row, 11]).Interior.Color = entry.Value.Item2;
                }
                else
                {
                    objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, 11], objxls.worksheet.Cells[n_row, 11]).Interior.Color = entry.Value.Item2;
                }

                //RECENTLY ADDED
                objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, 12], objxls.worksheet.Cells[n_row, 15]).Font.Bold = true;
                objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, 12], objxls.worksheet.Cells[n_row, 15]).Font.Size = 14;
                //objxls.worksheet.get_Range(objxls.worksheet.Cells[4, 12], objxls.worksheet.Cells[count_rows, 14]).Font.Color = 0x1f8a13; //#138a1f
                //objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, 12], objxls.worksheet.Cells[n_row, 15]).Font.Color = 0x0000ff; //#138a1f
                                                                                                                                        //objxls.worksheet.get_Range(objxls.worksheet.Cells[3, 8], objxls.worksheet.Cells[count_rows, 8]).Interior.Color = 0xC07000;
                objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, 8], objxls.worksheet.Cells[n_row, 8]).Interior.Color = 0xD58D53;
                //objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, 9], objxls.worksheet.Cells[n_row, 10]).Interior.Color = 0x00C0FF;
                objxls.worksheet.get_Range(objxls.worksheet.Cells[n_row, 9], objxls.worksheet.Cells[n_row, 10]).Interior.Color = 0x37B5FB;

                n_row++;
            }

            //objxls.worksheet.get_Range(objxls.worksheet.Cells[(n_row + 1), 1], objxls.worksheet.Cells[200, 16]).Value = "";
            objxls.worksheet.get_Range(objxls.worksheet.Cells[(n_row + 1), 1], objxls.worksheet.Cells[200, 16]).Value = "";
            //objxls.worksheet.get_Range(objxls.worksheet.Cells[(n_row + 1), 8], objxls.worksheet.Cells[200, 14]).Interior.Color = 0xffffff;
            //objxls.worksheet.get_Range(objxls.worksheet.Cells[(n_row + 1), 16], objxls.worksheet.Cells[200, 16]).Interior.Color = 0xffffff;

            objxls.worksheet.get_Range(objxls.worksheet.Cells[3, 1], objxls.worksheet.Cells[3, 16]).AutoFilter(1);

            for (int i = objxls.workbook.Worksheets.Count; i > 0; i--) //busca y borra la hoja shtCmb
            {
                Worksheet wkSheet = (Worksheet)objxls.workbook.Worksheets[i];
                if (wkSheet.Name == shtCombine)
                {
                    objxls.app.DisplayAlerts = false;
                    wkSheet.Delete();
                    objxls.app.DisplayAlerts = true;
                }
            }

            objxls.SelectSheet("Duplicate");

            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 2d;
            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 2d;
            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 2d;
            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 2d;
            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeTop].Color = 0x000000;
            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;

            ((Range)objxls.worksheet.get_Range("A200:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3d;
            ((Range)objxls.worksheet.get_Range("A3:P3")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3d;
            ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 4d;
            ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 4d;
            ((Range)objxls.worksheet.get_Range("O200:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 4d;

            ((Range)objxls.worksheet.get_Range("A200:P200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
            ((Range)objxls.worksheet.get_Range("A3:P3")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;
            ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = 0x000000;
            ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = 0x000000;
            ((Range)objxls.worksheet.get_Range("O200:O200")).Cells.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = 0x000000;

            ((Range)objxls.worksheet.get_Range("O4:O200")).Cells.Interior.Color = 0xE6E6E6;

            ((Range)objxls.worksheet.get_Range("A4:A200")).Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            ((Range)objxls.worksheet.get_Range("H4:O200")).Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            ((Range)objxls.worksheet.get_Range("A4:P200")).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;

            //((Range)objxls.worksheet.get_Range("A3:O3")).Cells.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 2d;

            //wkSheet.Cells[7, 1].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 1d;
            //xlWorkSheet5.Cells[7, 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = 1d;
            //xlWorkSheet5.Cells[7, 1].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 1d;
            //xlWorkSheet5.Cells[7, 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 1d;

            objxls.SelectSheet("Duplicate");

            if (originalSht.Length > 28)
            {
                originalSht = originalSht.Substring(0, 28);
            }
            else
            {
                originalSht = originalSht.ToString();
            }

            objxls.worksheet.Name = originalSht;

            return objxls;
        }

        private void bt_total_Click(object sender, EventArgs e)
        {
            this.st_label.Text = "Proceso para generar hoja de totales...";
            try
            {
                this.mensaje.Update();
                if (!System.IO.File.Exists(System.IO.Path.Combine(this.tb_path.Text, this.tb_save.Text)))
                {
                    this.picbox_save.Image = new Bitmap(Properties.Resources.uncheck, new Size(24, 24));
                    this.st_label.Text = "El archivo no existe...";
                    this.mensaje.Update();
                    MessageBox.Show(this.st_label.Text);
                    return;
                }
                this.st_label.Text = "Conectando con el archivo excel...";
                this.mensaje.Update();
            }
            catch
            {
                this.st_label.Text = "Error al buscar el archivo de Excel";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
                return;
            }
            ExcelInterface objxls = null;
            try
            {
                objxls = new ExcelInterface();
                if (!objxls.FindDocument(System.IO.Path.Combine(this.tb_path.Text, this.tb_save.Text)))
                {
                    this.st_label.Text = "Error al conectarse con el archivo de Excel";
                    this.mensaje.Update();
                    MessageBox.Show(this.st_label.Text);
                    return;
                }
            }
            catch
            {
                this.st_label.Text = "Error al conectarse con el archivo de Excel";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
                objxls = null;
                return;
            }
            //objxls.OpenFileToEdit(System.IO.Path.Combine(this.tb_path.Text,this.tb_save.Text));
            if (!objxls.SelectSheet("Total"))
            {
                this.st_label.Text = "Error: El archivo excel no contiene una hoja denominada Total";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
                objxls.ShowExcel();
                objxls.Dispose();
                objxls = null;
                return;
            }

            // HASTA ACA ESTA TODO BIEN


            //objxls.ShowExcel();
            Dictionary<string, int> l_data = new Dictionary<string, int>();
            Dictionary<string, string> l_resp = new Dictionary<string, string>();
            //objxls.worksheet.Select();
            objxls.HideExcel();

            int countaux = 1;
            string straux = string.Empty;
            long row, col;
            int quantity;
            int hojas;
            this.st_label.Text = "Tomando información...";
            this.mensaje.Update();
            hojas = 0;
            try
            {
                foreach (Excel.Worksheet wksht in objxls.workbook.Worksheets)
                {
                    if ((wksht.Name != "Total") && (wksht.Name != "Caratula") && (wksht.Name != "Referencia") && (wksht.Name != "Materiales") && (wksht.Name != "ComponentData"))
                    {
                        hojas++;
                        this.st_label.Text = "Tomando información " + (hojas) + "/" + (objxls.workbook.Worksheets.Count - 3) + " ...";
                        this.mensaje.Update();
                        objxls.worksheet = wksht;
                        if (objxls.worksheet != null)
                        {
                            countaux = 1;
                            row = 3;
                            try
                            {
                                var n_cant_tip = (objxls.worksheet.Cells[1, 5] as Excel.Range).Text;
                                countaux = Convert.ToInt32(n_cant_tip);
                            }
                            catch
                            {
                                countaux = 1;
                            }
                            while ((objxls.worksheet.Cells[row + 1, 1] as Excel.Range).Text.ToString() != string.Empty)
                            {
                                row++;
                                try
                                {
                                    //var n_cant_item = (objxls.worksheet.Cells[row, col_need] as Excel.Range).Text;
                                    var n_cant_item = (objxls.worksheet.Cells[row, col_need] as Excel.Range).Text;
                                    if (Convert.ToString(n_cant_item) == "")
                                        continue;
                                    //MessageBox.Show("Cant item: " + n_cant_item);
                                    quantity = Convert.ToInt32(n_cant_item);
                                }
                                catch
                                {
                                    MessageBox.Show("Hay un dato de cantidad no numérico en la hoja " + objxls.worksheet.Name + " Celda [" + row + "|" + col_need + "], se colocó como valor cantidad 1");
                                    quantity = 1;

                                }
                                quantity *= countaux;
                                straux = "";
                                try
                                {
                                    for (col = col_desc; col <= col_codefab; col++)
                                    {
                                        var str_data = (objxls.worksheet.Cells[row, col] as Excel.Range).Text ?? string.Empty;
                                        straux += str_data.ToString().Replace(';', ',') + ";";
                                    }
                                    if (!l_data.ContainsKey(straux))
                                    {
                                        l_data.Add(straux, quantity);
                                    }
                                    else
                                    {
                                        quantity += l_data[straux];
                                        l_data[straux] = quantity;
                                    }

                                    //Responsable
                                    var str_resp = (objxls.worksheet.Cells[row, col_resp] as Excel.Range).Text ?? string.Empty;
                                    if (!l_resp.ContainsKey(straux))
                                    {
                                        l_resp.Add(straux, str_resp.ToString());
                                    }
                                    else if (l_resp[straux] == "")
                                    {
                                        l_resp[straux] = str_resp.ToString();
                                    }
                                    else
                                    {
                                    }
                                }
                                catch
                                {
                                    this.st_label.Text = "Ocurrió un error al leer una fila de materiales en la Hoja " + objxls.worksheet.Name + " Celda [" + row + "|" + col_desc + "], se colocó como valor cantidad 1";
                                    this.mensaje.Update();
                                    MessageBox.Show(this.st_label.Text);
                                    objxls.ShowExcel();
                                    objxls.Dispose();
                                    objxls = null;
                                    return;
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                this.st_label.Text = "Error al intentar obtener datos de hojas";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
                objxls.ShowExcel();
                objxls.Dispose();
                objxls = null;
                return;
            }

            this.st_label.Text = "Conteo finalizado, completando hoja de totales...";
            this.mensaje.Update();

            try
            {
                Excel.Range c1 = null, c2 = null;
                if (objxls.SelectSheet("Total"))
                {
                    c1 = (Excel.Range)objxls.worksheet.Cells[2, 1];
                    c2 = (Excel.Range)objxls.worksheet.Cells[1000, 9];
                    //objxls.range = (Excel.Range)objxls.worksheet.get_Range("A2", "I1000");
                    try
                    {
                        objxls.range = (Excel.Range)objxls.worksheet.get_Range(c1, c2);
                        objxls.range.ClearContents();
                    }
                    catch
                    {
                        objxls.range = (Excel.Range)objxls.worksheet.Range[c1, c2];
                        objxls.range.ClearContents();
                    }
                    //objxls.range = (Excel.Range)objxls.worksheet.get_Range("K2", "M1000");
                    c1 = (Excel.Range)objxls.worksheet.Cells[2, 11];
                    c2 = (Excel.Range)objxls.worksheet.Cells[1000, 13];
                    try
                    {
                        objxls.range = (Excel.Range)objxls.worksheet.get_Range(c1, c2);
                        objxls.range.ClearContents();
                    }
                    catch
                    {
                        objxls.range = (Excel.Range)objxls.worksheet.Range[c1, c2];
                        objxls.range.ClearContents();
                    }
                }
            }
            catch
            {
                this.st_label.Text = "Error al intentar borrar los datos antiguos de la hoja de Total";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
                objxls.ShowExcel();
                objxls.Dispose();
                objxls = null;
                return;
            }

            //while ((objxls.worksheet.Cells[4, 1] as Excel.Range).Text.ToString() != string.Empty)
            //{
            //    (objxls.worksheet.Cells[4, 1] as Excel.Range).EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            //}
            long l_row = 2;
            int counter = 1;
            List<string> l_list = new List<string>();
            foreach (var item in l_data.Keys)
            {
                string s_aux = item;
                l_list.Add(s_aux);
            }
            try
            {
                foreach (var item in l_list.OrderBy(d => d))
                {
                    string s_aux = item;
                    //objxls.WriteValue(l_row, col_saldo, "=RC[-1]-RC[-2]");
                    try
                    {
                        objxls.range = objxls.worksheet.get_Range("A" + l_row.ToString() + "", "N" + l_row.ToString() + "");
                    }
                    catch (Exception)
                    {
                        objxls.range = objxls.worksheet.Range["A" + l_row.ToString() + "", "N" + l_row.ToString() + ""];
                    }
                    objxls.range.Font.Size = 9;

                    //objxls.range.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    //objxls.range.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    //objxls.range.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium;
                    //objxls.range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    //objxls.range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                    objxls.WriteValue(l_row, col_item, counter.ToString());
                    objxls.WriteValue(l_row, col_desc, s_aux);
                    objxls.WriteValue(l_row, col_need_total, l_data[s_aux].ToString());
                    objxls.WriteValue(l_row, col_build, "0");
                    objxls.WriteValue(l_row, col_resp, l_resp[s_aux].ToString());

                    //objxls.range = objxls.worksheet.get_Range("A" + l_row.ToString() + "", "A" + l_row.ToString() + "");
                    //objxls.range.Formula = objxls.range.Value;
                    //objxls.range = objxls.worksheet.get_Range("H" + l_row.ToString() + "", "J" + l_row.ToString() + "");
                    //objxls.range.Formula = objxls.range.Value;
                    //objxls.range = objxls.worksheet.get_Range("M" + l_row.ToString() + "", "N" + l_row.ToString() + "");
                    //objxls.range.Formula = objxls.range.Value;

                    l_row++;
                    counter++;

                }
            }
            catch
            {
                this.st_label.Text = "Error al completar la hoja de Total";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
                objxls.Dispose();
                objxls = null;
                return;
            }

            l_row--;
            l_list = null;
            l_data = null;

            try
            {
                objxls.worksheet.PageSetup.PrintArea = "$A$1:$N$" + l_row.ToString() + "";
                this.st_label.Text = "Guardando archivo...";
                this.mensaje.Update();
                objxls.workbook.RefreshAll();
                objxls.workbook.Save();
                objxls.SelectSheet("Referencia");
                objxls.SelectSheet("Total");
            }
            catch
            {
                this.st_label.Text = "Error al intentar guardar el archivo";
                this.mensaje.Update();
                MessageBox.Show(this.st_label.Text);
                objxls.ShowExcel();
                objxls.Dispose();
                objxls = null;
                return;
            }
            this.st_label.Text = "Conteo completado";
            this.mensaje.Update();
            MessageBox.Show(this.st_label.Text);
            objxls.ShowExcel();
            objxls.Dispose();
            objxls = null;
        }

        private void tb_path_TextChanged(object sender, EventArgs e)
        {
            if (System.IO.Directory.Exists(this.tb_path.Text))
                this.tb_path.BackColor = System.Drawing.SystemColors.Window;
            else
                this.tb_path.BackColor = System.Drawing.Color.IndianRed;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.st_label.Text = "";
            this.mensaje.Update();

            string route = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);

            string excelfilepath = route.Split(new string[] { "bin" }, StringSplitOptions.None)[0] + "/bin/Release/Empty_BillOfMaterial.xlsm";

            this.Hide();
            var formAddShts = new AddAdditionalShts();

            formAddShts.Closed += (s, args) =>
            {
                this.Show();

                string[] values = { };

                if (formAddShts.DocValues != "")
                {
                    values = formAddShts.DocValues.Split('\n');
                }

                bool firstRowAdded = false;

                if (dataGridView1.RowCount < 1)
                {
                    dataGridView1.Rows.Add();
                    firstRowAdded = true;
                }

                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = Regex.Replace(values[i], @"\p{C}+", string.Empty);

                    List<string> addedFiles = new List<string>();

                    foreach (var item in this.dataGridView1.Rows)
                    {
                        addedFiles.Add((string)((DataGridViewRow)item).Cells[0].Value);
                    }

                    DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                    row.Cells[0].Value = values[i];
                    row.Cells[1].Value = "Hoja vacía";
                    row.Cells[3].Value = excelfilepath;

                    if (!(addedFiles.Contains(values[i])))
                    {
                        this.dataGridView1.Rows.Add(row);
                        this.dataGridView1.Rows[this.dataGridView1.Rows.Count - 1].Cells[2].ReadOnly = true;
                        //tmp_data = null;
                        // icon = null;
                        this.st_label.Text = "Documento agregado a la lista";
                        this.mensaje.Update();
                    }
                    else
                    {
                        MessageBox.Show("No se puede incluir el archivo " + values[i] + " porque ya existe ese nombre.", "Agregar archivo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                if (firstRowAdded)
                {
                    dataGridView1.Rows.RemoveAt(0);
                }

                List<string> nameFiles = new List<string>();

                nameFiles.Add("");

                foreach (var item in this.dataGridView1.Rows)
                {
                    if ((string)((DataGridViewRow)item).Cells[1].Value != "Hoja vacía")
                    {
                        if ((string)((DataGridViewRow)item).Cells[0].Value == "BoMList")
                        {
                            nameFiles.Add((string)((DataGridViewRow)item).Cells[1].Value);
                        }
                        else
                        {
                            nameFiles.Add((string)((DataGridViewRow)item).Cells[0].Value);
                        }
                    }
                }

                for (int j = 0; j < this.dataGridView1.Rows.Count; j++)
                {
                    string combineWith = (string)(((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2].Value);
                    ((DataGridViewComboBoxCell)((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2]).Items.Clear();

                    foreach (var option in nameFiles)
                    {
                        if ((option != (((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[0].Value).ToString()) &&
                            ((option != (((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[1].Value).ToString()) ||
                            (((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[0].Value).ToString() != "BoMList"))
                        {
                            ((DataGridViewComboBoxCell)((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2]).Items.Add(option);
                        }
                    }

                    if (nameFiles.Contains(combineWith))
                    {
                        ((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2].Value = combineWith;
                    }
                    else
                    {
                        ((DataGridViewRow)this.dataGridView1.Rows[j]).Cells[2].Value = "";
                    }
                }

            };

            formAddShts.Show();

        }

        private void Generate_GiveFeedback(object sender, GiveFeedbackEventArgs e)
        {

        }

        private void Generate_FormClosing(object sender, FormClosingEventArgs e)
        {
            app.Quit();
        }
    }
}
