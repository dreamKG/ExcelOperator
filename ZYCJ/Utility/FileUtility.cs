using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZYCJ.Utility
{
    public static class FileUtility
    {
        public static void CopyFile(string sourceDirectory, string fileName, string targetFilePath)
        {
            DirectoryInfo di = new DirectoryInfo(sourceDirectory);
            FileInfo[] files = di.GetFiles($"*{fileName}*");
            if (files.Length == 0)
                throw new Exception($"未在{sourceDirectory}下找到匹配特征字符串为*{fileName}*的文件，");
            FileInfo sourceFile = files[0];

            FileInfo targetFile = new FileInfo(targetFilePath);

            if (File.Exists(targetFilePath))
            {
                targetFile.Delete();
            }
            //sourceFile.CopyTo(targetFilePath,false);
            File.Copy(sourceFile.FullName, targetFilePath, true);
        }

        public static string GetFileNameExtension(string sourceDirectory, string fileName)
        {
            DirectoryInfo di = new DirectoryInfo(sourceDirectory);
            FileInfo[] files = di.GetFiles($"*{fileName}*");
            if (files.Length == 0)
                throw new Exception($"未在{sourceDirectory}下找到匹配特征字符串为*{fileName}*的文件，");
            FileInfo sourceFile = files[0];

            string fullName = sourceFile.FullName;

            string[] temp = fullName.Split('.');
            int tempCount = temp.Length - 1;
            string fileNameExtension = temp[tempCount];

            return fileNameExtension;
        }

    }
}
