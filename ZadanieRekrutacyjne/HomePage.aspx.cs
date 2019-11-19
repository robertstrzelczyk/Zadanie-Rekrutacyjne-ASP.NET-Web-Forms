using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ZadanieRekrutacyjneRS
{
    public partial class HomePage : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Create_Click(object sender, EventArgs e)
        {
            try
            {
                string PathFile = "C:\\rekrutacja_backend\\robert_strzelczyk.xlsx";

                if (File.Exists(PathFile))
                {

                    File.Delete(PathFile);
                }

                Excel.Application oApplication;
                Excel.Worksheet oWorksheet;
                Excel.Workbook oWorkbooks;

                oApplication = new Excel.Application();
                oWorkbooks = oApplication.Workbooks.Add();
                oWorksheet = (Excel.Worksheet)oWorkbooks.Worksheets.get_Item(1);
                oWorksheet.Cells[1, 1] = "ID";
                oWorksheet.Cells[1, 2] = "Imie Nazwisko";
                oWorksheet.Cells[1, 3] = "Kwota";

                string[,] NamesTables = new string[4, 3];

                NamesTables[0, 0] = "1";
                NamesTables[0, 1] = "Adam Nowak";
                NamesTables[0, 2] = "200.00 zł";
                NamesTables[1, 0] = "2";
                NamesTables[1, 1] = "Jan Kowalski";
                NamesTables[1, 2] = "150.20 zł";
                NamesTables[2, 0] = "3";
                NamesTables[2, 1] = "Jan Hak";
                NamesTables[2, 2] = "50.00 zł";
                NamesTables[3, 0] = "4";
                NamesTables[3, 1] = "Henryk Kania";
                NamesTables[3, 2] = "2000.00 zł";

                oWorksheet.get_Range("A2", "C5").Value2 = NamesTables;
                oWorksheet.get_Range("A1", "C1").Font.Bold = true;
                oWorksheet.get_Range("B1", "B100").ColumnWidth = 20;
                oWorksheet.get_Range("C1", "C100").ColumnWidth = 20;


                oWorkbooks.SaveAs(PathFile);
                oWorkbooks.Close();
                Response.Write("<script>alert('Plik został utworzony pomyślnie')</script>");
                oApplication.Quit();
            }
            catch (Exception theException)
            {
                Response.Write("<script>alert('Plik nie został utworzony')</script>");
            }
        }
    }
}