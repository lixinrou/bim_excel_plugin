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

namespace ExcelAddIn
{
    public partial class QueryConfigPanel : Form
    {
        public QueryConfigPanel()
        {
            InitializeComponent();
            BuildDataSource();
        }

        private bool update = false;

        private void skinButton2_Click(object sender, EventArgs e)
        {
            string appsetting = ConfigurationManager.ConnectionStrings["mysql"].ToString();
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
            CCWin.MessageBoxEx.Show(message);
        }

        private void skinButton1_Click(object sender, EventArgs e)
        {
            string assemblyConfigFile = Assembly.GetExecutingAssembly().Location;
            string appDomainConfigFile = AppDomain.CurrentDomain.SetupInformation.ConfigurationFile;

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);


            //获取appSettings节点
            AppSettingsSection appSettings = (AppSettingsSection)config.GetSection("appSettings");

            //获取连接串节点
            ConnectionStringsSection connectionSettings = (ConnectionStringsSection)config.GetSection("connectionStrings");
            connectionSettings.ConnectionStrings.Remove(skinTextBox4.Text);
            connectionSettings.ConnectionStrings.Add(new ConnectionStringSettings
            {
                Name = skinTextBox4.Text,
                ConnectionString = $"server={skinTextBox3.Text};User Id={skinTextBox1.Text};password={skinTextBox2.Text};Database={skinTextBox5.Text}",
                ProviderName = "MySql.Data.MySqlClient"
            });
            update = true;
            //删除name，然后添加新值
            appSettings.Settings.Remove("name");
            appSettings.Settings.Add("name", "user");

            //保存配置文件
            config.Save();
            ConfigurationManager.RefreshSection("connectionStrings");
            ConfigurationManager.RefreshSection("appSettings");
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
    }
}
