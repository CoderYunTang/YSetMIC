using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace YSetMIC
{
    public static class RegistryHelper
    {
        /// <summary>
        /// 读取指定名称的注册表的值
        /// </summary>
        /// <param name="root">注册表根:Registry.LocalMachine</param>
        /// <param name="subkey">子项节点名称:SOFTWARE\\Microsoft</param>
        /// <param name="name">要获取的节点名称</param>
        /// <returns>注册表值</returns>
        public static string GetRegistryValue(RegistryKey root, string subkey, string name)
        {

            using (RegistryKey myKey = root.OpenSubKey(subkey, true))
            {
                string registData = "";
                if (myKey != null)
                {
                    registData = myKey.GetValue(name).ToString();
                }

                return registData;
            }
        }

        /// <summary>
        /// 写注册表数据
        /// </summary>
        /// <param name="root">注册表根:Registry.LocalMachine</param>
        /// <param name="subkey">子项节点名称:SOFTWARE\\Microsoft</param>
        /// <param name="name">要写入的节点名称</param>
        /// <param name="value">要写入的值</param>
        /// <param name="registryValueKind">写入的值类型</param>
        public static void SetRegistryValue(RegistryKey root, string subkey, string name, object value, RegistryValueKind registryValueKind)
        {
            using (RegistryKey aimdir = root.OpenSubKey(subkey,RegistryKeyPermissionCheck.ReadWriteSubTree, System.Security.AccessControl.RegistryRights.SetValue))
            {
                if (aimdir != null)
                {
                    aimdir.SetValue(name, value, registryValueKind);
                }
            }
        }
    }
}
