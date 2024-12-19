using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ListOfMaterialsPGGA.Forms
{
    public partial class AddAdditionalShts : Form
    {
        public AddAdditionalShts()
        {
            InitializeComponent();

            this.dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            //this.dataGridView1.Columns[0].Width = 500;

            int i = 100;
            while (i > 0)
            {
                this.dataGridView1.Rows.Add();
                i--;
            }
        }

        string docValues = "";

        public string DocValues
        {
            get { return docValues; }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            docValues = "";
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            docValues = "";

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if ((row.Cells[0].Value != null) && (row.Cells[0].Value != DBNull.Value) && !(String.IsNullOrWhiteSpace(row.Cells[0].Value.ToString())) &&
                    (row.Cells[0].Value.ToString() != ""))
                {
                    docValues += row.Cells[0].Value.ToString() + "\n";
                }
                else
                {
                    if(docValues.Length > 0) {
                        docValues = docValues.Remove(docValues.Length - 1);
                    }
                    break;
                }
            }


            //docValues = textBox2.Text;
            this.Close();
        }

            //        if(e.KeyCode == Keys.V)
            //{
            //    MessageBox.Show("Key V was pressed");

            //    if (e.Control)
            //    {
            //        MessageBox.Show("There was a paste command executed");

            //        string s = Clipboard.GetText();
            //        string[] lines = s.Split('\n');
            //        int row = dataGridView1.CurrentCell.RowIndex;
            //        int col = dataGridView1.CurrentCell.ColumnIndex;
            //        foreach (string line in lines)
            //        {
            //            string[] cells = line.Split('\t');
            //            int cellsSelected = cells.Length;
            //            if (row < dataGridView1.Rows.Count)
            //            {
            //                for (int i = 0; i < cellsSelected; i++)
            //                {
            //                    if (col + i < dataGridView1.Columns.Count)
            //                        dataGridView1[col + i, row].Value = cells[i];
            //                    else
            //                        break;
            //                }
            //                row++;
            //            }
            //            else
            //            {
            //                break;
            //            }

            //            //if ((row == (dataGridView1.Rows.Count - 1)) || (row == dataGridView1.Rows.Count))
            //            //{
            //            //    dataGridView1.Rows.Add();
            //            //}
            //        }
            //    }
            //}

        private void keyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.V)
            {
                if (e.Control)
                {
                    string s = Clipboard.GetText();
                    string[] lines = s.Split('\n');
                    int row = dataGridView1.CurrentCell.RowIndex;
                    int col = dataGridView1.CurrentCell.ColumnIndex;
                    foreach (string line in lines)
                    {
                        string[] cells = line.Split('\t');
                        int cellsSelected = cells.Length;
                        if (row < dataGridView1.Rows.Count)
                        {
                            for (int i = 0; i < cellsSelected; i++)
                            {
                                if (col + i < dataGridView1.Columns.Count)
                                    dataGridView1[col + i, row].Value = cells[i];
                                else
                                    break;
                            }
                            row++;
                        }
                        else
                        {
                            break;
                        }

                        //if ((row == (dataGridView1.Rows.Count - 1)) || (row == dataGridView1.Rows.Count))
                        //{
                        //    dataGridView1.Rows.Add();
                        //}
                    }
                }
            }
        }

        private void finish_edit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dgv = this.dataGridView1;

            if((dgv.Rows[e.RowIndex].Cells[0].Value != null) && (dgv.Rows[e.RowIndex].Cells[0].Value != DBNull.Value) && 
                !(String.IsNullOrWhiteSpace((dgv.Rows[e.RowIndex].Cells[0].Value).ToString())))
            {
                string route = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
                dgv.Rows[e.RowIndex].Cells[1].Value = route.Split(new string[] { "bin" }, StringSplitOptions.None)[0] + "/bin/Release/Empty_BillOfMaterial.xlsm";
            } else
            {
                dgv.Rows[e.RowIndex].Cells[1].Value = "";
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void bt_up_Click(object sender, EventArgs e)
        {
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
    }
}
