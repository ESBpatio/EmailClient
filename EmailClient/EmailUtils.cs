using System;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace EmailClient
{
    public class EmailUtils
    {
        public void sendMessage(string emailClient, string serverAddress
            , int port, string login, string password , string emailAdress
            , string displayName, string bodyMessage, string subject
            )
        {
            MailAddress from = new MailAddress(emailAdress,displayName);
            MailAddress to = new MailAddress(emailClient);
            MailMessage mailMessage = new MailMessage(from, to);

            mailMessage.Subject = subject;
            mailMessage.Body = bodyMessage; //string.Format("<h2>Произошла ошибка при загрузке письма. Ошибка : {0}</h2>", error);
            mailMessage.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient(serverAddress, 587);
            smtp.Credentials = new NetworkCredential(login, password);
            smtp.EnableSsl = true;
            smtp.Send(mailMessage);
        }
        public void sendMessageToAttacj(
            string emailClient, string error, string serverAddress
            , int port, string login, string password,string emailAdress
            , string displayName, string bodyMessage, string subject, string fileName
            )
        {
            MailAddress from = new MailAddress(emailAdress, displayName); //"info.price@patio-minsk.by", "ESBinfo");
            MailAddress to = new MailAddress(emailClient);
            MailMessage mailMessage = new MailMessage(from, to);

            mailMessage.Subject = subject; //"Ошибка при загрузке вложения";
            mailMessage.Body = bodyMessage; //string.Format("<h2>Произошла ошибка при загрузке письма. Ошибка : {0}</h2>", error);
            mailMessage.IsBodyHtml = true;

            mailMessage.Attachments.Add(new Attachment(fileName));
            SmtpClient smtp = new SmtpClient(serverAddress, 587);
            smtp.Credentials = new NetworkCredential(login, password);
            smtp.EnableSsl = true;
            smtp.Send(mailMessage);
        }

        public bool AddFileToDist(MemoryStream memoryStream, string fileName, string patchToDisk)
        {
            try
            {
                DirectoryInfo dirInfo = new DirectoryInfo(patchToDisk);
                //Создаем каталог для файла
                if (!dirInfo.Exists)
                    dirInfo.Create();

                using (FileStream fs = new FileStream(patchToDisk + fileName, FileMode.OpenOrCreate))
                {
                    memoryStream.WriteTo(fs);
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }
    }
}
