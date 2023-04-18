using ExcelAddIn.Entity;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Common
{
    public class Settings
    {
        private AppSettingsSection appSettings;
        private ConnectionStringsSection connectionSettings;
        private Configuration config;
        private readonly string APPSETTINGS = "appSettings";
        private readonly string CONNECTIONSTRINGS = "connectionStrings";

        /// <summary>
        /// 构造函数
        /// </summary>
        public Settings()
        {
            config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            //获取appSettings节点
            appSettings = (AppSettingsSection)config.GetSection(APPSETTINGS);

            //获取连接串节点
            connectionSettings = (ConnectionStringsSection)config.GetSection(CONNECTIONSTRINGS);
        }

        /// <summary>
        /// 创建数据库连接
        /// </summary>
        /// <param name="connectionEntity"></param>
        public void AddConnection(ConnectionEntity connectionEntity)
        {           
            connectionSettings.ConnectionStrings.Add(new ConnectionStringSettings
            {
                Name = connectionEntity.ConnectionName,
                ConnectionString = $"server={connectionEntity.Host};User Id={connectionEntity.UserName};password={connectionEntity.Password};Database={connectionEntity.Database}",
                ProviderName = connectionEntity.Provider
            });
            //保存配置文件
            Save(CONNECTIONSTRINGS);
        }

        /// <summary>
        /// 移除数据库连接
        /// </summary>
        /// <param name="name"></param>
        public void RemoveConnection(string name)
        {
            connectionSettings.ConnectionStrings.Remove(name);
            Save(CONNECTIONSTRINGS);
        }

        /// <summary>
        /// 添加配置节
        /// </summary>
        /// <param name="settingEntity"></param>
        public void AddSettings(SettingEntity settingEntity)
        {
            appSettings.Settings.Add(settingEntity.key, settingEntity.value);
            Save(APPSETTINGS);
        }

        /// <summary>
        /// 移除指定的配置节
        /// </summary>
        /// <param name="key"></param>
        public void RemoveSettings(string key)
        {
            appSettings.Settings.Remove(key);
            Save(APPSETTINGS);
        }

        /// <summary>
        /// 获取配置项
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public string GetSettingValue(string key)
        {
            return appSettings.Settings[key].Value;
        }

        /// <summary>
        /// 保存配置文件
        /// </summary>
        /// <param name="section"></param>
        private void Save(string section)
        {
            config.Save();
            ConfigurationManager.RefreshSection(section);
        }
    }
}
