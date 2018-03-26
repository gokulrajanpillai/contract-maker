
namespace ContractApplikation.Src.Controller
{
    class Class1
    {
        public static void WordFromExcel(string filepath)
        {
            string wbkName = filepath;

            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook workbook = xlApp.Workbooks.Open(wbkName);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Sheets[1];

            Microsoft.Office.Interop.Word._Application wdApp = new Microsoft.Office.Interop.Word.Application();
            wdApp.Visible = true;
            Microsoft.Office.Interop.Word.Document document = wdApp.Documents.Add();

            worksheet.Range["A1", "G7"].Copy();
            document.Range().PasteSpecial();

            workbook.Close();
            xlApp.Quit();
        }
    }
}
