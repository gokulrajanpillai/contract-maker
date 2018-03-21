
using ContractApplikation.Src.Model;
using Spire.Doc;
using Spire.Doc.Documents;
using System.IO;
using System.Windows.Forms;

namespace ContractApplikation.Src.Controller
{
    public class DocumentManager
    {
        private static FileFormat DocumentFormat = FileFormat.Docx;

        private static string CurrentProjectPath()
        {
            return Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
        }

        private static string MasterDocumentPath()
        {
            return CurrentProjectPath() + "\\vertrag.docx";
        }

        private static string SaveDocumentPath()
        {
            return CurrentProjectPath();
        }

        private static string FinishedContractDocumentPath()
        {
            return CurrentProjectPath();
        }

        private static Document LoadDocument(string documentFilePath)
        {
            Document doc = new Document();
            doc.LoadFromFile(documentFilePath);
            return doc;
        }

        private static void SaveDocument(Document doc, string nameOfDocument)
        {
            doc.SaveToFile(SaveDocumentPath() + "\\" + nameOfDocument, DocumentFormat);
        }

        public static void CreateSampleDocument()
        {
            Document doc = new Document();
            Section section = doc.AddSection();
            Paragraph para = section.AddParagraph();
            para.AppendText("Created my first document!");

            MessageBox.Show("Directory: "+Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName +"\\CreatedWordDocument.docx");
            doc.SaveToFile(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName +"\\CreatedWordDocument.docx", FileFormat.Docx);
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
            Document doc = LoadDocument(MasterDocumentPath());

            foreach (Section section in doc.Sections)
            {
                foreach (Paragraph para in section.Paragraphs)
                {
                    string paragraph = para.Text;
                    
                    foreach (string word in paragraph.Split(' '))
                    {
                        switch (word)
                        {
                            case "[Kunden_Anrede]"          : paragraph = paragraph.Replace("[Kunden_Anrede]", Kunden.Anrede); break;
                            case "[Kunden_Vorname]"         : paragraph = paragraph.Replace("[Kunden_Vorname]", Kunden.Vorname); break;
                            case "[Kunden_Nachname]"        : paragraph = paragraph.Replace("[Kunden_Nachname]", Kunden.Nachname); break;
                            case "[Kunden_Vollname]"        : paragraph = paragraph.Replace("[Kunden_Vollname]", Kunden.Name); break;
                            case "[Kunden_Firma]"           : paragraph = paragraph.Replace("[Kunden_Firma]", Kunden.Firma); break;
                            case "[Kunden_Geschäftsbereich]": paragraph = paragraph.Replace("[Kunden_Geschäftsbereich]", Kunden.Geschäftsbereich); break;
                            case "[Kunden_Abteilungszusatz]": paragraph = paragraph.Replace("[Kunden_Abteilungszusatz]", Kunden.Abteilungszusatz); break;
                            case "[Kunden_Abteilung]"       : paragraph = paragraph.Replace("[Kunden_Abteilung]", Kunden.Abteilung); break;
                            case "[Kunden_Email]"           : paragraph = paragraph.Replace("[Kunden_Email]", Kunden.Email); break;
                            case "[Kunden_Telefon]"         : paragraph = paragraph.Replace("[Kunden_Telefon]", Kunden.Telefon); break;
                            case "[Kunden_Strasse]"         : paragraph = paragraph.Replace("[Kunden_Strasse]", Kunden.Strasse); break;
                            case "[Kunden_PLZ]"             : paragraph = paragraph.Replace("[Kunden_PLZ]", Kunden.PLZ); break;
                            case "[Kunden_Ort]"             : paragraph = paragraph.Replace("[Kunden_Ort]", Kunden.Ort); break;

                            case "[Projekt_Projektnummer]"      : paragraph = paragraph.Replace("[Projekt_Projektnummer]", Projekt.Projektnummer); break;
                            case "[Projekt_StartDatum]"         : paragraph = paragraph.Replace("[Projekt_StartDatum]", Projekt.StartDatum); break;
                            case "[Projekt_EndDatum]"           : paragraph = paragraph.Replace("[Projekt_EndDatum]", Projekt.EndDatum); break;
                            case "[Projekt_AnzahlStunden]"      : paragraph = paragraph.Replace("[Projekt_AnzahlStunden]", Projekt.AnzahlStunden); break;
                            case "[Projekt_Verrechnungssatz]"   : paragraph = paragraph.Replace("[Projekt_Verrechnungssatz]", Projekt.Verrechnungssatz); break;
                            case "[Projekt_Einzelpreis]"        : paragraph = paragraph.Replace("[Projekt_Einzelpreis]", Projekt.Einzelpreis); break;
                            case "[Projekt_AngebotSumme]"       : paragraph = paragraph.Replace("[Projekt_AngebotSumme]", Projekt.AngebotSumme); break;
                            case "[Projekt_ProjektTitel]"       : paragraph = paragraph.Replace("[Projekt_ProjektTitel]", Projekt.ProjektTitel); break;
                            case "[Projekt_Koordinator]"        : paragraph = paragraph.Replace("[Projekt_Koordinator]", Projekt.Koordinator); break;
                            case "[Projekt_Gesprächsperson]"    : paragraph = paragraph.Replace("[Projekt_Gesprächsperson]", Projekt.Gesprächsperson); break;
                            case "[Projekt_Disponent]"          : paragraph = paragraph.Replace("[Projekt_Disponent]", Projekt.Disponent); break;
                            case "[Projekt_ProjektBeschreibung]": paragraph = paragraph.Replace("[Projekt_ProjektBeschreibung]", Projekt.ProjektBeschreibung); break;
                        }
                    }
                    para.Text = paragraph;
                }
            }
            doc.SaveToFile(FinishedContractDocumentPath() + "\\" + NameOfDocument, DocumentFormat);
            MessageBox.Show("File processed and saved successfully");
        }
    }
}
