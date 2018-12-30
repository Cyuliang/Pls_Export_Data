using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pls_Export_Data
{
    public partial class Form1 : Form
    {

        public static Form1 _Form;
        public Form1()
        {
            InitializeComponent();

            _Form = this;

            dateTimePicker1.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.ShowUpDown = true;
            dateTimePicker1.Value = DateTime.Now.Date;
            dateTimePicker2.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.ShowUpDown = true;
            dateTimePicker2.Value = DateTime.Now.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

            TimeradioButton.Checked = true;
        }

        /// <summary>
        /// 查询数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Findbutton_Click(object sender, EventArgs e)
        {
            try
            {
                dataSet1 = FindData();
                dataGridView1.DataSource = dataSet1.Tables[0].DefaultView;
                bindingSource1.DataSource = dataGridView1.DataSource;
                bindingNavigator1.BindingSource = bindingSource1;
                dataGridView1.Columns[4].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";
                dataGridView1.Columns[5].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";
            }
            catch (Exception)
            {
                MessageBox.Show("访问远程数据库失败，请重新查询一次");
            }
        }

        /// <summary>
        /// 导出数据
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Exportbutton_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                string saveFileName = "";
                //bool fileSaved = false;  
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    DefaultExt = "xls",
                    Filter = "Excel文件|*.xls",
                    FileName = string.Format("InData_{0:yyyyMMddHHmmss}.xls", DateTime.Now)
                };
                saveDialog.ShowDialog();
                saveFileName = saveDialog.FileName;
                if (saveFileName.IndexOf(":") < 0)
                {
                    return; //被点了取消   
                }
                else
                {
                    var t1 = new Task(TaskMethod, saveFileName, TaskCreationOptions.LongRunning);
                    t1.Start();
                }
            }
            else
            {
                MessageBox.Show("报表为空,无表格需要导出", "提示", MessageBoxButtons.OK);
            }
        }

        /// <summary>
        /// 查询数据
        /// </summary>
        /// <param name="cmdText"></param>
        /// <returns></returns>
        private DataSet FindData()
        {
            MySqlParameter[] parameter = {
                new MySqlParameter("@DataS",MySqlDbType.DateTime),
                new MySqlParameter("@DataE",MySqlDbType.DateTime),
                new MySqlParameter("@Plate",MySqlDbType.VarChar)
            };

            parameter[0].Value = dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss");
            parameter[1].Value = dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm:ss");
            parameter[2].Value = PlateText.Text;

            string cmdText = string.Empty;

            if (TimeradioButton.Checked)
            {
                cmdText = "select * from hw.rundata where InDatetime between @DataS and @DataE";
            }
            if (DataradioButton.Checked)
            {
                if (PlateText.Text.Trim() != string.Empty)
                {
                    cmdText = "select * from hw.rundata where Plate=@Plate";
                }
                else
                {
                    MessageBox.Show("请输入要查询的车牌");
                    PlateText.Focus();
                }
            }

            return MySqlHelper.GetDataSet(MySqlHelper.Conn, CommandType.Text, cmdText, parameter);
        }

        static object taskMethodLocj = new object();
        static void TaskMethod(object title)
        {
            lock (taskMethodLocj)
            {
                DataGridView dataGridView1 = _Form.dataGridView1;

                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (xlApp == null)
                {
                    MessageBox.Show("无法创建Excel对象，可能您的机子未安装Excel");
                    return;
                }
                Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1  

                //写入标题  
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                }
                //写入数值  
                for (int r = 0; r < dataGridView1.Rows.Count; r++)
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (i == 7)
                        {
                            worksheet.Cells[r + 2, i + 1] = "'" + dataGridView1.Rows[r].Cells[i].Value;
                        }
                        else
                        {
                            worksheet.Cells[r + 2, i + 1] = dataGridView1.Rows[r].Cells[i].Value;
                        }
                    }
                    System.Windows.Forms.Application.DoEvents();
                }
                worksheet.Columns.EntireColumn.AutoFit();//列宽自适应  
                                                         //   if (Microsoft.Office.Interop.cmbxType.Text != "Notification")  
                                                         //   {  
                                                         //       Excel.Range rg = worksheet.get_Range(worksheet.Cells[2, 2], worksheet.Cells[ds.Tables[0].Rows.Count + 1, 2]);  
                                                         //      rg.NumberFormat = "00000000";  
                                                         //   }  

                if (title.ToString() != "")
                {
                    try
                    {
                        workbook.Saved = true;
                        workbook.SaveCopyAs(title.ToString());
                        //fileSaved = true;  
                    }
                    catch (Exception ex)
                    {
                        //fileSaved = false;  
                        MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);
                    }

                }
                //else  
                //{  
                //    fileSaved = false;  
                //}  
                xlApp.Quit();
                GC.Collect();//强行销毁   
                             // if (fileSaved && System.IO.File.Exists(saveFileName)) System.Diagnostics.Process.Start(saveFileName); //打开EXCEL  
                MessageBox.Show("导出文件成功", "提示", MessageBoxButtons.OK);
            }
        }
    }
}
