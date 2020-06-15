using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;
using System.Drawing;

namespace Phd
{
    class CPhdConfig
    {
        private string m_strFilePath;

        public CPhdConfig(string strFilePath)
        {
            m_strFilePath = strFilePath;
        }

        /// <summary>
        /// 写ini文件函数
        /// </summary>
        /// <param name="Section">Section</param>
        /// <param name="Key">关键字</param>
        /// <param name="Value">要设置的值</param>
        /// <param name="filepath">ini文件路径</param>
        public void IniWriteValue(string Section, string Key, string Value)//对ini文件进行写操作的函数
        {
            WritePrivateProfileString(Section, Key, Value, m_strFilePath);
        }

        /// <summary>
        /// 读ini文件函数
        /// </summary>
        /// <param name="Section">Section</param>
        /// <param name="Key">关键字</param>
        /// <param name="filepath">文件路径</param>
        /// <returns>返回string</returns>
        public string IniReadValue(string Section, string Key)//对ini文件进行读操作的函数
        {
            StringBuilder temp = new StringBuilder(255);
            int i = GetPrivateProfileString(Section, Key, "", temp, 255, m_strFilePath);
            return temp.ToString();
        }


        #region DLL导入
        //函数作用：向INI文件中写信息
        [DllImport("kernel32")]
        public static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        //函数作用：从私有文件中获取字符串（读Ini文件）
        //section:欲在其中查找条目的小节名称
        //key:欲获取的项名或条目名
        //def:指定的条目没有找到时返回的默认值。可设为空（""） 
        //retVal:指定一个字串缓冲区，长度至少为size 
        //size:缓冲区大小
        //filePath:INI文件的完整路径
        [DllImport("kernel32")]
        public static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        #endregion
    }
}
