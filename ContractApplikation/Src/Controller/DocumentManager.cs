
using ContractApplikation.Src.Helper;
using ContractApplikation.Src.Model;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;
using System.Windows.Forms;

namespace ContractApplikation.Src.Controller
{
    public class DocumentManager
    {
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

            foreach (Section section in doc.Sections)
            {
                foreach (Paragraph para in section.Paragraphs)
                {
                    string paragraph = para.Text;
                    paragraph = ReplaceCustomerPlaceholders(paragraph, Kunden);
                    paragraph = ReplaceProjektPlaceholders(paragraph, Projekt);
                    para.Text = paragraph;
                }
            }

            SaveDocument(doc, NameOfDocument);
            MessageBox.Show("File processed and saved successfully");
        }

        private static string ReplaceCustomerPlaceholders(string paragraph, Ansprechpartner kunden)
        {
            paragraph = paragraph.Replace("[Kunden_Anrede]", kunden.Anrede);
            paragraph = paragraph.Replace("[Kunden_Vorname]", kunden.Vorname);
            paragraph = paragraph.Replace("[Kunden_Nachname]", kunden.Nachname);
            paragraph = paragraph.Replace("[Kunden_Vollname]", kunden.Name);
            paragraph = paragraph.Replace("[Kunden_Firma]", kunden.Firma);
            paragraph = paragraph.Replace("[Kunden_Geschäftsbereich]", kunden.Geschäftsbereich);
            paragraph = paragraph.Replace("[Kunden_Abteilungszusatz]", kunden.Abteilungszusatz);
            paragraph = paragraph.Replace("[Kunden_Abteilung]", kunden.Abteilung);
            paragraph = paragraph.Replace("[Kunden_Email]", kunden.Email);
            paragraph = paragraph.Replace("[Kunden_Telefon]", kunden.Telefon);
            paragraph = paragraph.Replace("[Kunden_Strasse]", kunden.Strasse);
            paragraph = paragraph.Replace("[Kunden_PLZ]", kunden.PLZ);
            paragraph = paragraph.Replace("[Kunden_Ort]", kunden.Ort);

            return paragraph;
        }


        private static string ReplaceProjektPlaceholders(string paragraph, Projekt project)
        {
            paragraph = paragraph.Replace("[Projekt_Projektnummer]", project.Projektnummer);
            paragraph = paragraph.Replace("[Projekt_StartDatum]", project.StartDatum);
            paragraph = paragraph.Replace("[Projekt_EndDatum]", project.EndDatum);
            paragraph = paragraph.Replace("[Projekt_AnzahlStunden]", project.AnzahlStunden.ToString());
            paragraph = paragraph.Replace("[Projekt_Verrechnungssatz]", project.Verrechnungssatz.ToString());
            paragraph = paragraph.Replace("[Projekt_Einzelpreis]", project.Einzelpreis);
            paragraph = paragraph.Replace("[Projekt_AngebotSumme]", project.AngebotSumme);
            paragraph = paragraph.Replace("[Projekt_ProjektTitel]", project.ProjektTitel);
            paragraph = paragraph.Replace("[Projekt_Koordinator]", project.Koordinator);
            paragraph = paragraph.Replace("[Projekt_Gesprächsperson]", project.Gesprächsperson);
            paragraph = paragraph.Replace("[Projekt_Disponent]", project.Disponent);
            paragraph = paragraph.Replace("[Projekt_ProjektBeschreibung]", project.ProjektBeschreibung);

            return paragraph;
        }
    }
}
