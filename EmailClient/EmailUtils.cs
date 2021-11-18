using MailKit.Net.Imap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace EmailClient
{
    internal class EmailUtils
    {
        public void sendMessage(string emailClient, string error , string serverAddress , int port,string login, string password)
        {
            MailAddress from = new MailAddress("info.price@patio-minsk.by", "ESBinfo");
            MailAddress to = new MailAddress(emailClient);
            MailMessage mailMessage = new MailMessage(from, to);
            
            mailMessage.Subject = "Ошибка при загрузке вложения";
            mailMessage.Body = string.Format("<h2>Произошла ошибка при загрузке письма. Ошибка : {0}</h2>", error);
            mailMessage.IsBodyHtml = true;
            SmtpClient smtp = new SmtpClient(serverAddress, 587);
            smtp.Credentials = new NetworkCredential(login, password);
            smtp.EnableSsl = true;
            smtp.Send(mailMessage);
        }
    }
}
