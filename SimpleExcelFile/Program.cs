using System;

namespace SimpleExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Creating Excel File and send Email!");

            try
            {
                string fromEmail = "";
                string fromName = "";
                string password = "";
                string toEmail = "email1@contoso.com,email2@contoso.com,email3@contoso.com";
                string filename = "";
                string subject = "";
                string body = "";

                Email email = new Email(fromEmail, fromName, password, toEmail, filename, subject, body);
                ExcelFile ef = new ExcelFile();
                ef.SetDate(DateTime.Now);
                ef.SetHeaders();
                ef.SaveAs(filename);

                email.SendEmail();


                Console.WriteLine("Excel File en Email send successful!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel File and Email cannot send, because of: " + ex.Message);
                Console.Write("Press any key to continue...");
                Console.ReadKey();
            }

        }
    }
}
