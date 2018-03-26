
using ContractApplikation.Src.Helper;
using ContractApplikation.Src.Model;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Xls;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ContractApplikation.Src.Controller
{
    public class DocumentManager
    {
        public static bool includeCostTable = true;

        private static readonly Spire.Doc.FileFormat DocumentFormat = Spire.Doc.FileFormat.Docx;

        private static string PrototypeDocumentPath()
        {
            return Constants.FileLocation.PROTOTYPE_CONTRACT;
        }

        private static Document LoadDocument(string documentFilePath)
        {
            Document doc = new Document();
            doc.LoadFromFile(documentFilePath);
            return doc;
        }

        private static void SaveDocument(Document doc, string NameOfDocument)
        {
            doc.SaveToFile(Constants.FileLocation.OutputFilePath(NameOfDocument), DocumentFormat);
        }

        public static void CreateSampleDocument()
        {
            Document doc = new Document();
            Section section = doc.AddSection();
            Paragraph para = section.AddParagraph();
            para.AppendText("Created my first document!");
            MessageBox.Show("Directory: " + Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "\\CreatedWordDocument.docx");
            doc.SaveToFile(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "\\CreatedWordDocument.docx", DocumentFormat);
        }

        public static void DisplayDocumentWithName(string name)
        {
            DisplayDocumentWithPath(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "\\" + name);
        }



        public static void DisplayDocumentWithPath(string path)
        {
            Document doc = new Document();
            doc.LoadFromFile(path);

            foreach (Section section in doc.Sections)
            {
                MessageBox.Show("Section: " + section.ToString());
                foreach (Paragraph para in section.Paragraphs)
                {
                    MessageBox.Show("Paragraph: " + para.Text);
                }
            }
        }

        public static void GenerateContractDocument(string NameOfDocument, Ansprechpartner Kunden, Projekt Projekt)
        {
            Document doc = LoadDocument(PrototypeDocumentPath());
            ReplaceCustomerPlaceholders(ref doc, Kunden);
            ReplaceProjektPlaceholders(ref doc, Projekt);
            AddProjectCostTable(ref doc, Projekt);

            SaveDocument(doc, NameOfDocument);
            MessageBox.Show("File processed and saved successfully");
            OpenDocument(NameOfDocument);
        }

        private static void AddProjectCostTable(ref Document doc, Projekt projekt)
        {
            // Load the workbook in the WebBrowser control
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(Constants.FileLocation.OutputFilePath(projekt.CostTableFileName));
            Worksheet sheet = workbook.Worksheets[0];

            Section section = doc.Sections[0];
            TextSelection selection = doc.FindString("[Projekt_TabelleKosten]", true, true);
            TextRange range = selection.GetAsOneRange();
            Paragraph paragraph = range.OwnerParagraph;
            Body body = paragraph.OwnerTextBody;
            int index = body.ChildObjects.IndexOf(paragraph);

            Table table = section.AddTable(true);
            table.ResetCells(sheet.LastRow, sheet.LastColumn);

            //Traverse the rows and columns of table in worksheet and get the cells, call a custom function CopyStyle() to copy the font style and cell style from Excel to Word table. 
            for (int r = 1; r <= sheet.LastRow; r++)
            {
                for (int c = 1; c <= sheet.LastColumn; c++)
                {
                    CellRange xCell = sheet.Range[r, c];
                    TableCell wCell = table.Rows[r - 1].Cells[c - 1];

                    //Fill data to Word table 
                    TextRange textRange = wCell.AddParagraph().AppendText(xCell.NumberText);

                    //Copy the formatting of table to Word 
                    CopyStyle(textRange, xCell, wCell);

                }
            }
            //Set column width of Word table in Word 
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Rows[i].Cells.Count; j++)
                {
                    table.Rows[i].Cells[j].Width = 100f;
                }
            }
            table.IndentFromLeft = paragraph.Format.LeftIndent;

            body.ChildObjects.Remove(paragraph);
            body.ChildObjects.Insert(index, table);
        }

        //The custom function CopyStyle() is defined as below 
        private static void CopyStyle(TextRange wTextRange, CellRange xCell, TableCell wCell)
        {
            //Copy font style 
            wTextRange.CharacterFormat.TextColor = xCell.Style.Font.Color;
            wTextRange.CharacterFormat.FontSize = (float)xCell.Style.Font.Size;
            wTextRange.CharacterFormat.FontName = xCell.Style.Font.FontName;
            wTextRange.CharacterFormat.Bold = xCell.Style.Font.IsBold;
            wTextRange.CharacterFormat.Italic = xCell.Style.Font.IsItalic;

            //Copy backcolor 
            wCell.CellFormat.BackColor = xCell.Style.Color;

            //Copy text alignment 
            switch (xCell.HorizontalAlignment)
            {
                case HorizontalAlignType.Left:
                    wTextRange.OwnerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                    break;
                case HorizontalAlignType.Center:
                    wTextRange.OwnerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    break;
                case HorizontalAlignType.Right:
                    wTextRange.OwnerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    break;
            }
        }

        private static bool FileExistsInOutputDirectory(string NameOfDocument)
        {
            return File.Exists(Constants.FileLocation.OutputFilePath(NameOfDocument));
        }

        private static void OpenDocument(string NameOfDocument)
        {
            Process.Start(Constants.FileLocation.OutputFilePath(NameOfDocument));
        }

        private static void ReplaceCustomerPlaceholders(ref Document doc, Ansprechpartner kunden)
        {
            doc.Replace("[Kunden_Anrede]", kunden.Anrede, true, false);

            doc.Replace("[Kunden_Vorname]", kunden.Vorname, true, false);
            doc.Replace("[Kunden_Vollname]", kunden.Name, true, false);
            doc.Replace("[Kunden_Nachname]", kunden.Nachname, true, false);
            doc.Replace("[Kunden_Firma]", kunden.Firma, true, false);
            doc.Replace("[Kunden_Geschäftsbereich]", kunden.Geschäftsbereich, true, false);
            doc.Replace("[Kunden_Abteilungszusatz]", kunden.Abteilungszusatz, true, false);
            doc.Replace("[Kunden_Abteilung]", kunden.Abteilung, true, false);
            doc.Replace("[Kunden_Email]", kunden.Email, true, false);
            doc.Replace("[Kunden_Telefon]", kunden.Telefon, true, false);
            doc.Replace("[Kunden_Strasse]", kunden.Strasse, true, false);
            doc.Replace("[Kunden_PLZ]", kunden.PLZ, true, false);
            doc.Replace("[Kunden_Ort]", kunden.Ort, true, false);
        }


        private static void ReplaceProjektPlaceholders(ref Document doc, Projekt project)
        {
            if (!includeCostTable)
                doc.Replace("[Projekt_TabelleKosten]", "", true, false);

            doc.Replace("[Projekt_ProjektTitel]", project.ProjektTitel, true, true);
            doc.Replace("[Projekt_Projektnummer]", project.Projektnummer, true, false);
            doc.Replace("[Projekt_StartDatum]", project.StartDatum, true, false);
            doc.Replace("[Projekt_EndDatum]", project.EndDatum, true, false);
            doc.Replace("[Projekt_AnzahlStunden]", project.AnzahlStunden.ToString(), true, false);
            doc.Replace("[Projekt_Verrechnungssatz]", project.Verrechnungssatz.ToString(), true, false);
            doc.Replace("[Projekt_Einzelpreis]", project.Einzelpreis, true, false);
            doc.Replace("[Projekt_AngebotSumme]", project.AngebotSumme, true, false);
            doc.Replace("[Projekt_Koordinator]", project.Koordinator, true, false);
            doc.Replace("[Projekt_Gesprächsperson]", project.Gesprächsperson, true, false);
            doc.Replace("[Projekt_Disponent]", project.Disponent, true, false);
            doc.Replace("[Projekt_ProjektBeschreibung]", project.ProjektBeschreibung, true, false);
        }

        internal static void EditCostTableForProject(Projekt projekt)
        {
            // Copy file from prototype, if it does not exist
            if (!FileExistsInOutputDirectory(projekt.CostTableFileName))
                File.Copy(Constants.FileLocation.PROTOTYPE_COSTTABLE, Constants.FileLocation.OutputFilePath(projekt.CostTableFileName));
            OpenDocument(projekt.CostTableFileName);
        }

        private static void CopyExcelFile(string sourcePath, string destinationPath)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(sourcePath);
            
            workbook.SaveToFile(destinationPath, ExcelVersion.Version97to2003);
            System.Diagnostics.Process.Start(destinationPath);

        }
    }
}
