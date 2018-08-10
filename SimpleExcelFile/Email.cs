using System;
using System.Net;
using System.Net.Mail;

namespace SimpleExcelFile
{
    class Email
    {
        private string _fromEmail;
        private string _fromName;
        private string _password;
        private string _toEmail;
        private string _filename;
        private string _subject;
        private string _body;
        public Email(string fromEmail, string fromName, string password, string toEmail, string filename, string subject, string body)
        {
            _fromEmail = fromEmail;
            _fromName = fromName;
            _password = password;
            _toEmail = toEmail;
            _filename = filename;
            _subject = subject;
            _body = body;
        }

        public void SendEmail()
        {
            try
            {
                var fromAddress = new MailAddress(_fromEmail, _fromName);

                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, _password)
                };
                var message = new MailMessage();
                message.From = fromAddress;
                message.Subject = _subject;
                message.Body = _body;
                message.To.Add(_toEmail);
                Attachment _attach = new Attachment(_filename);
                message.Attachments.Add(_attach);
                smtp.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error send email, " + ex.Message);
                Console.Write("Press any key to continue...");
                Console.ReadKey();
            }
        }
    }
}
