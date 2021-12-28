using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Web;

namespace WagnerCable.Utilities
{
    public class UtilityManager
    {
        public static bool SendEmail(string toAddress, string subject, string content)
        {
            try
            {
                SmtpClient smtp = new SmtpClient();
                smtp.Port = 587;
                smtp.Host = "smtp.yandex.com.tr";
                smtp.EnableSsl = true;
                smtp.Credentials = new NetworkCredential("mikayilsadigzade@lumhar.com", "aa122300");

                MailMessage mail = new MailMessage();
                mail.From = new MailAddress("mikayilsadigzade@lumhar.com", subject);
                mail.To.Add(toAddress);
                mail.Subject = subject.Replace("\r\n", "");
                mail.IsBodyHtml = true;
                mail.Body = content;
                smtp.Send(mail);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }
    }
}