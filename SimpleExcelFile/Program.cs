using System;

namespace SimpleExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Creating Excel File!");

            try
            {
                string filename = "Reporte_Horas.xlsx";
                
                Email email = new Email("[fromEmail]", "[fromName]", "[password]", "[toEmail]", "[filename]", "[subject]", "[body]");
                ExcelFile ef = new ExcelFile();
                ef.SetDate(DateTime.Now);
                //ef.SetDate(new DateTime(2018, 7, 31));
                ef.SetHeaders();
                ef.SaveAs(filename);

                email.SendEmail();


                Console.WriteLine("Excel File created successful!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Excel File cannot create, because of: " + ex.Message);
                Console.Write("Press any key to continue...");
                Console.ReadKey();
            }

        }
    }
}
