using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailClient
{
    public static class fileUtils
    {
        public static async Task WriteSetting(byte[] body, string pathCatalog, string pathFile ,string format)
        {
            //Создания каталога для файлов
            DirectoryInfo directoryInfo = new DirectoryInfo(pathCatalog);
            FileInfo fileInfo = new FileInfo(pathFile);
            if (!directoryInfo.Exists)
                directoryInfo.Create();
            else
            {
                directoryInfo.Create();
            }
                
            if(fileInfo.Exists)
                fileInfo.Delete();
            using (FileStream fstream = new FileStream($"{pathFile}",FileMode.OpenOrCreate))
            {
                await fstream.WriteAsync(body, 0, body.Length);
            }
        }
        public static string GetSetting(string pathFile)
        {
            using (StreamReader sr = new StreamReader(pathFile))
            {
                return sr.ReadToEnd();
            }
        }
    }
}
