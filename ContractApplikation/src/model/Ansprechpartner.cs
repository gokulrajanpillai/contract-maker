using ContractApplikation.Src.Helper;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ContractApplikation.Src.Model
{

    public enum Honorifics
    {
        HERR,
        FRAU
    }

    public class Ansprechpartner
    {
        public String Bezeichnung { get; private set; }

        public String Vorname { get; private set; }

        public String Nachname { get; private set; }

        public String Abteilung { get; private set; }

        public String Email { get; private set; }

        public String Telefon { get; private set; }

        public String Strasse { get; private set; }

        public String Ort { get; private set; }

        public String Firma { get; private set; }

        public String Abteilungszusatz { get; private set; }

        public String Geschäftsbereich { get; private set; }

        // Custom property: not part of the database
        public String Name
        {
            get
            {
                return Utilities.FirstLetterToUpperCase(Bezeichnung) + ". " + Utilities.FirstLetterToUpperCase(Vorname) + " " + Utilities.FirstLetterToUpperCase(Nachname);
            }
        }
        public Ansprechpartner(List<TextBox> listOfTextboxes, Honorifics bezeichnung)
        {
            this.Bezeichnung = bezeichnung.ToString();

            foreach (TextBox textBox in listOfTextboxes)
            {
                this.GetType().GetProperty(textBox.Name).SetValue(this, textBox.Text);
            }
        }

        public Ansprechpartner(OleDbDataReader dataReader)
        {
            this.Bezeichnung         = dataReader.GetValue(1).ToString();
            this.Vorname             = dataReader.GetValue(2).ToString();
            this.Nachname            = dataReader.GetValue(3).ToString();
            this.Abteilung           = dataReader.GetValue(4).ToString();
            this.Email               = dataReader.GetValue(5).ToString();
            this.Telefon             = dataReader.GetValue(6).ToString();
            this.Strasse             = dataReader.GetValue(7).ToString();
            this.Ort                 = dataReader.GetValue(8).ToString();
            this.Firma               = dataReader.GetValue(9).ToString();
            this.Abteilungszusatz    = dataReader.GetValue(10).ToString();
            this.Geschäftsbereich    = dataReader.GetValue(11).ToString();
        }

    }
}
