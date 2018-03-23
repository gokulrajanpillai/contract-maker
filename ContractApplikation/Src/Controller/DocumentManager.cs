
using ContractApplikation.Src.Helper;
using ContractApplikation.Src.Model;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ContractApplikation.Src.Controller
{
    public class DocumentManager
    {
        public static bool includeCostTable  = false;

        private static readonly FileFormat DocumentFormat = FileFormat.Docx;

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
            doc.SaveToFile(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + "\\CreatedWordDocument.docx", FileFormat.Docx);
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
            SaveDocument(doc, NameOfDocument);
            MessageBox.Show("File processed and saved successfully");
            OpenDocument(NameOfDocument);
        }

        private static void OpenDocument(string NameOfDocument)
        {
            Process.Start(Constants.FileLocation.OutputFilePath(NameOfDocument));
        }

        private static void ReplaceCustomerPlaceholders(ref Document doc, Ansprechpartner kunden)
        {
            doc.Replace("[Kunden_Anrede]", kunden.Anrede, true, false);
            doc.Replace("[Kunden_Vorname]", kunden.Vorname, true, false);
            doc.Replace("[Kunden_Nachname]", kunden.Nachname, true, false);
            doc.Replace("[Kunden_Vollname]", kunden.Name, true, false);
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
    }
}
