using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
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
            //else
            //{
            //    directoryInfo.Create();
            //}
                
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
        public static DirectoryInfo CreateCatalog(string pathCatalog)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(pathCatalog);
            if (!directoryInfo.Exists)
                directoryInfo.Create();
            return directoryInfo;
        }
        public static FileInfo DowloadFileFromURLToPath(string url , string path)
        {
            try
            {
                using (WebClient webClient = new WebClient())
                {
                    webClient.DownloadFile(url, path);
                    return new FileInfo(path);
                }
            }
            catch (Exception)
            {

                return null;
            }
        }
        public static MemoryStream GetFileStream(string path)
        {
            MemoryStream ms = new MemoryStream();
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            byte[] bytes = new byte[file.Length];
            file.Read(bytes, 0, (int)file.Length);
            ms.Write(bytes, 0, (int)file.Length);
            return ms;
        }
    }
}
