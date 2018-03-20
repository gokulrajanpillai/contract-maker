using ContractApplikation.Src.Helper;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ContractApplikation.Src.Model
{

    public enum Salutation
    {
        HERR,
        FRAU
    }

    public class Ansprechpartner
    {
        public String Anrede { get; private set; }

        public String Vorname { get; private set; }

        public String Nachname { get; private set; }

        public String Abteilung { get; private set; }

        public String Email { get; private set; }

        public String Telefon { get; private set; }

        public String Strasse { get; private set; }

        public String PLZ { get; private set; }

        public String Ort { get; private set; }

        public String Firma { get; private set; }

        public String Abteilungszusatz { get; private set; }

        public String Geschäftsbereich { get; private set; }

        // Custom property: not part of the database
        public String Name
        {
            get
            {
                return Utilities.FirstLetterToUpperCase(Anrede.ToLower()) + ". " + Utilities.FirstLetterToUpperCase(Vorname) + " " + Utilities.FirstLetterToUpperCase(Nachname);
            }
        }
        public Ansprechpartner(List<TextBox> listOfTextboxes, Salutation bezeichnung)
        {
            this.Anrede = bezeichnung.ToString();

            foreach (TextBox textBox in listOfTextboxes)
            {
                this.GetType().GetProperty(Utilities.FirstLetterToUpperCase(textBox.Name)).SetValue(this, textBox.Text);
            }
        }

        public Ansprechpartner(OleDbDataReader dataReader)
        {
            this.Anrede              = dataReader.GetValue(1).ToString();
            this.Vorname             = dataReader.GetValue(2).ToString();
            this.Nachname            = dataReader.GetValue(3).ToString();
            this.Abteilung           = dataReader.GetValue(4).ToString();
            this.Email               = dataReader.GetValue(5).ToString();
            this.Telefon             = dataReader.GetValue(6).ToString();
            this.Strasse             = dataReader.GetValue(7).ToString();
            this.PLZ                 = dataReader.GetValue(8).ToString();
            this.Ort                 = dataReader.GetValue(9).ToString();
            this.Firma               = dataReader.GetValue(10).ToString();
            this.Abteilungszusatz    = dataReader.GetValue(11).ToString();
            this.Geschäftsbereich    = dataReader.GetValue(12).ToString();
        }

    }
}
