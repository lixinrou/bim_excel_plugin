using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Configuration;
using CCWin.SkinControl;
using System.Reflection;
using CCWin;
using ExcelAddIn.Common;

namespace ExcelAddIn
{
    public partial class ConfigPanel : Form
    {
        public ConfigPanel()
        {
            InitializeComponent();
            BuildDataSource();
        }

        private bool update = false;

        private void skinButton2_Click(object sender, EventArgs e)
        {
            Settings settings = new Settings();
            string current = settings.GetSettingValue("active");
            string appsetting = ConfigurationManager.ConnectionStrings[current].ToString();
            MySqlConnection conn = new MySqlConnection(appsetting);
            string message = string.Empty;
            try
            {
                Console.WriteLine(conn.ConnectionTimeout);
                conn.Open();
                string state = conn.State.ToString();
                message = state == "Open" ? "测试连接成功！" : "测试连接失败！";
            }
            catch (Exception ex)
            {
                message = "测试连接失败！";
            }
            MessageBoxEx.Show(message);
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {            
            Settings settings = new Settings();
            settings.RemoveConnection(skinTextBox4.Text);
            settings.AddConnection(new Entity.ConnectionEntity
            {
                ConnectionName = skinTextBox4.Text,
                Host = skinTextBox3.Text,
                UserName = skinTextBox1.Text,
                Password = skinTextBox2.Text,
                Database = skinTextBox5.Text
            });            
            update = true;            
        }

        private void skinComboBox1_Click(object sender, EventArgs e)
        {
            if (update == true)
            {
                BuildDataSource();
                update = false;
            }
        }

        private void BuildDataSource()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            ConnectionStringsSection connectionSettings = (ConnectionStringsSection)config.GetSection("connectionStrings");
            List<string> dataSource = new List<string>();

            int count = connectionSettings.ConnectionStrings.Count;
            for (int i = 0; i < count; i++)
            {
                dataSource.Add(connectionSettings.ConnectionStrings[i].Name);
            }
            skinComboBox1.DataSource = dataSource;
        }

        private void skinButton4_Click(object sender, EventArgs e)
        {
            string currentConnection = skinComboBox1.SelectedItem.ToString();
            Settings settings = new Settings();
            Entity.SettingEntity settingEntity = new Entity.SettingEntity();
            settingEntity.value = currentConnection;
            settings.RemoveSettings(settingEntity.key);
            settings.AddSettings(settingEntity);
            MessageBoxEx.Show($"连接[{currentConnection}]已激活");
        }
    }
}
