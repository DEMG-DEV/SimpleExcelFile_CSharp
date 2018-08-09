using ClosedXML.Excel;
using System;
using System.Globalization;
using System.Threading;

namespace SimpleExcelFile
{
    class ExcelFile
    {
        private XLWorkbook _wb;
        private readonly string[] _cells = { "A", "B", "C", "D", "E", "F" };
        private readonly string _workHours = "9.5";
        private DateTime _date;
        DateTime _begining;
        DateTime _end;

        public ExcelFile()
        {
            _wb = new XLWorkbook();
            Thread.CurrentThread.CurrentCulture = new CultureInfo("es-ES");
        }

        public void SetDate(DateTime date)
        {
            _date = date;
        }

        public void SaveAs(string filename)
        {
            _wb.SaveAs(filename);
        }

        public void SetHeaders()
        {
            #region Add Data
            var _ws = _wb.Worksheets.Add("Reporte de Horas");
            GetWeek(_date, new CultureInfo("es-ES"), out _begining, out _end);
            _ws.Cell("A1").Value = "Mes de " + _date.ToString("MMMM");
            _ws.Cell("A2").Value = "Semana de " + _begining.ToString("dd/MM/yyyy") + " al " + _end.AddDays(-1).ToString("dd/MM/yyyy");
            _ws.Cell("A4").Value = "Actividad";
            _ws.Cell("A5").Value = "Programación";

            foreach (string column in _cells)
            {
                if (!column.Equals("A"))
                {
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("es-ES");
                    _ws.Cell(column + "4").Value = _begining.ToString("dddd") + " " + _begining.Day;
                    _ws.Cell(column + "5").Value = _workHours;
                    _ws.Cell(column + "6").Value = _workHours;
                    _begining = _begining.AddDays(1);
                }
            }
            #endregion

            #region set Styles
            _ws.Style.Font.SetFontName("Calibri");
            _ws.Cell("A1").Style.Font.SetFontSize(22.0);
            _ws.Cell("A1").Style.Font.SetBold(true);
            _ws.Cell("A2").Style.Font.SetFontSize(16.0);
            _ws.Cell("A2").Style.Font.SetBold(true);
            foreach (string column in _cells)
            {
                _ws.Cell(column + "4").Style.Font.SetFontSize(18.0);
                _ws.Cell(column + "4").Style.Font.SetBold(true);
                _ws.Cell(column + "4").Style.Font.SetUnderline();
                _ws.Cell(column + "4").Style.Fill.BackgroundColor = XLColor.FromHtml("5b9bd5");


                _ws.Column("A").Width = 14;
                if (!column.Equals("A"))
                {
                    _ws.Column(column).AdjustToContents();
                    _ws.Cell(column + "6").Style.Fill.BackgroundColor = XLColor.FromHtml("ffc000");
                }
            }

            _ws.Row(1).AdjustToContents();
            _ws.Row(2).AdjustToContents();
            _ws.Row(3).AdjustToContents();
            _ws.Row(4).AdjustToContents();
            _ws.Row(5).AdjustToContents();
            _ws.Row(6).AdjustToContents();
            #endregion
        }

        private void GetWeek(DateTime now, CultureInfo cultureInfo, out DateTime begining, out DateTime end)
        {
            if (now == null)
                throw new ArgumentNullException("now");
            if (cultureInfo == null)
                throw new ArgumentNullException("cultureInfo");

            var firstDayOfWeek = cultureInfo.DateTimeFormat.FirstDayOfWeek;

            int offset = firstDayOfWeek - now.DayOfWeek;
            if (offset != 1)
            {
                DateTime weekStart = now.AddDays(offset);
                DateTime endOfWeek = weekStart.AddDays(5);
                begining = weekStart;
                end = endOfWeek;
            }
            else
            {
                begining = now.AddDays(-6);
                end = now;
            }
        }
    }
}
