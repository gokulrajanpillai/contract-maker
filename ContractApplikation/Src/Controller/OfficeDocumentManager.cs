
using ContractApplikation.Src.Helper;
using System.Diagnostics;
using System.IO;

namespace ContractApplikation.Src.Controller
{
    class OfficeDocumentManager
    {
        public static string WordFromExcel(string filepath)
        {
            killOfficeProcesses();

            Microsoft.Office.Interop.Excel._Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook workbook = xlApp.Workbooks.Open(filepath);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Sheets[1];
            Microsoft.Office.Interop.Word._Application wdApp = new Microsoft.Office.Interop.Word.Application();
            wdApp.Visible = false;

            Microsoft.Office.Interop.Word.Document document = wdApp.Documents.Add();
            worksheet.UsedRange.Copy();
            document.Range().PasteSpecial();

            if (File.Exists(Constants.FileLocation.COSTTABLE_WORDDOC))
                File.Delete(Constants.FileLocation.COSTTABLE_WORDDOC);
            document.SaveAs(Constants.FileLocation.COSTTABLE_WORDDOC);

            workbook.Close();
            xlApp.Workbooks.Close();
            xlApp.Quit();
            wdApp.Quit();

            killOfficeProcesses();

            return Constants.FileLocation.COSTTABLE_WORDDOC;
        }

        private static void killOfficeProcesses()
        {
            killExcelProcesses();
            killWordProcesses();
        }

        private static void killExcelProcesses()
        {
            foreach (Process excelProcess in Process.GetProcessesByName("EXCEL"))
            {
                excelProcess.Kill();
            }
        }

        private static void killWordProcesses()
        {
            foreach (Process excelProcess in Process.GetProcessesByName("WINWORD"))
            {
                excelProcess.Kill();
            }
        }
    }
}
