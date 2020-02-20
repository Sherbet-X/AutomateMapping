using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutomateMappingTool
{
    class DgvSettings
    {
        public void SetDgv(DataGridView dataGridView, string file, string sheetName, List<string> lstHeader)
        {
            Control.CheckForIllegalCrossThreadCalls = false;
            try
            {
                //connect excel
                if (System.IO.File.Exists(file))
                {
                    string connString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;""", file);
                    string query = string.Format("select * from [" + sheetName + "]", connString);
                    OleDbDataAdapter dtAdapter = new OleDbDataAdapter(query, connString);
                    DataSet ds = new DataSet();
                    dtAdapter.Fill(ds);

                    ////Remove header
                    //for (int i = 0; i < 2; i++)
                    //{
                    //    ds.Tables[0].Rows[i].Delete();
                    //}

                    ds.AcceptChanges();
                    ds.Tables[0].Rows.Add();
                    dataGridView.DataSource = ds.Tables[0];

                    //Set header
                    for (int i = 0; i < lstHeader.Count; i++)
                    {
                        dataGridView.Columns[i].HeaderText = lstHeader[i];
                    }

                    //Set style
                    dataGridView.ColumnHeadersDefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Bold);
                    dataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(100, 61, 167);
                    dataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    dataGridView.AutoResizeColumns();
                    dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    dataGridView.EnableHeadersVisualStyles = false;

                    //remove emtry row
                    for (int i = 0; i < dataGridView.RowCount; i++)
                    {
                        if (dataGridView.Rows[dataGridView.RowCount - 1].Cells[0].Value.ToString() == "")
                        {
                            dataGridView.Rows.RemoveAt(dataGridView.RowCount - 1);
                            i--;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot read data from file " + file + "\r\n" +
                    "Please check format file and field" + "\r\n" + ex.Message);
            }
        }
    }
}
