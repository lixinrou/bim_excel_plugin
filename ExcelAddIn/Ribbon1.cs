using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;//using Excel
using System.Windows.Forms;
using System.Configuration;
using CCWin.SkinControl;
using ExcelAddIn.Common;
using MySql.Data.MySqlClient;
using ExcelAddIn.Data;
using ExcelAddIn.Events;
using ExcelAddIn.Loading;
using System.Threading;
using System.Diagnostics;
using ExcelAddIn.Entity;

namespace ExcelAddIn
{
    public partial class Ribbon1
    {
        public event EventHandler SendMsgEvent;
        private ReferenceComponents refercomponents;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //TODO:加载初始化数据，如数据库连接
            //string appsetting = ConfigurationManager.AppSettings["excel"].ToString();
            refercomponents = new ReferenceComponents();
            SendMsgEvent += refercomponents.AfterParentFrmTextChange;

        }

        private void SingleCell_Click(object sender, RibbonControlEventArgs e)
        {
            //声明一个Excel的单元格区域变量。
            Excel.Range rang;

            //获得Excel所选区域
            rang = Globals.ThisAddIn.Application.Selection;

            if (rang.Count == 0) return;

            //所选区域底纹颜色设置为黄色。
            //rang.Interior.Color = System.Drawing.Color.Yellow;            
            int columns = rang.Columns.Count + 1;
            int rows = rang.Rows.Count + 1;

            // 遍历时先列后行，否则会出现矩阵反转
            for (int i = 1; i < columns; i++)
            {
                for (int k = 1; k < rows; k++)
                {
                    Excel.Range cell = rang.Item[i][k];
                    if (cell.Value2 == null)
                    {
                        continue;
                    }
                    string cellValue = ((string)cell.Value2).ToString();
                    if (!cellValue.Contains(','))
                    {
                        cell.Value2 = "'" + cellValue.Trim() + "'";
                    }
                    else
                    {
                        cellValue = cellValue.Replace(",", "','");
                        cell.Value2 = "'" + cellValue.Trim() + "'";
                    }

                }
            }

        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            ConfigPanel fm1 = new ConfigPanel();//创建新实例。
            fm1.ShowDialog();//以对话框的形式，在Excel中显示Form1。
        }

        private void MultiCell_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range rang;
            rang = Globals.ThisAddIn.Application.Selection;
            if (rang.Count == 0) return;
            int columns = rang.Columns.Count + 1;
            int rows = rang.Rows.Count + 1;
            StringBuilder builder = new StringBuilder();
            for (int i = 1; i < columns; i++)
            {
                for (int k = 1; k < rows; k++)
                {
                    Excel.Range cell = rang.Item[i][k];
                    cell.Value2 = "'" + cell;
                    if (cell.Value2 == null)
                    {
                        continue;
                    }
                    string cellValue = ((string)cell.Value2).ToString();
                    builder.Append(cellValue).Append(",");

                }
            }
            string cellsData = builder.ToString().TrimEnd(',');
            if (!cellsData.Contains(","))
            {
                cellsData = "'" + cellsData.Trim() + "'";
            }
            else
            {
                cellsData = cellsData.Replace(",", "','");
                cellsData = "'" + cellsData.Trim() + "'";
            }

            Form f = new Form();
            f.Width = f.Height = 300;
            f.Opacity = 0.9;
            f.StartPosition = FormStartPosition.CenterParent;
            SkinTextBox box = new SkinTextBox();
            box.Text = cellsData;
            box.ReadOnly = true;
            f.Controls.Add(box);
            f.ShowDialog();
        }

        private void System_Group_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void group_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            TextConfigPanel panel = new TextConfigPanel();
            panel.ShowDialog();
        }

        private void group2_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            QueryConfigPanel panel = new QueryConfigPanel();
            panel.ShowDialog();
        }

        

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range range;
            
            range = Globals.ThisAddIn.Application.Selection;
            if (range.Cells.Count != 1) return;
            Excel.Range cell = range.Item[1][1];
            string filter = (string)cell.Value2;
            if (string.IsNullOrEmpty(filter)) return;
            if (filter.Contains(","))
            {
                filter = filter.Replace(",", "','");
            }
            
            ParameterizedThreadStart thread = new ParameterizedThreadStart(UpdateReferenceComponents);
            
            LoadingHelper.ShowLoading("正在处理中，请稍候...", null , thread, new ReferencePipeline {
                objectId = "1681888161170880.23546",
                fileId = "1681888161170880",
                filterSystem = filter
            });
            refercomponents.ShowDialog();
        }
        

        private void UpdateReferenceComponents(object param)
        {
            ReferencePipeline pipeline = param as ReferencePipeline;

            //弹出新窗体，查询指定的系统下构件集合
            DataSource dataSource = new DataSource();
            MySqlConnection connection = (MySqlConnection)dataSource.GetCurrentConnection("mysql");
            StringBuilder builder = new StringBuilder();
            int count = 0;
            using (connection)
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(string.Format(ExcuteSql.SYSTEM_FILTER, pipeline.filterSystem, pipeline.fileId), connection);
                MySqlDataReader reader = cmd.ExecuteReader();
                builder.Append("[");
                while (reader.Read())
                {
                    object o = reader[0];
                    builder.Append("\"");
                    builder.Append(o.ToString()).Append("\",");
                    count++;
                }
                builder.Append("]");
                builder.Remove(builder.Length - 2, 1);
            }
            SendMsgEvent(this, new CustonEventArgs() { Text = builder.ToString(), ObjectId = pipeline.objectId ,Count = count});
        }
    }
}
