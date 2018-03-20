using ContractApplikation.Src.Helper;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ContractApplikation.Src.Model
{
    public class Projekt
    {
        public String Projektnummer { get; private set; }

        public String StartDatum { get; private set; }

        public String EndDatum { get; private set; }

        public String AnsprechpartnerID { get; private set; }

        public String AnzahlStunden { get; private set; }

        public String Verrechnungssatz { get; private set; }

        public String ProjektTitel { get; private set; }

        public String Gesprächsperson { get; private set; }

        public String Disponent { get; private set; }

        public String ProjektBeschreibung { get; private set; }

        public Projekt(List<TextBox> listOfTextboxes)
        {
            foreach (TextBox textBox in listOfTextboxes)
            {
                this.GetType().GetProperty(Utilities.FirstLetterToUpperCase(textBox.Name)).SetValue(this, textBox.Text);
            }
        }

        public Projekt(OleDbDataReader dataReader)
        {
            this.Projektnummer       = dataReader.GetValue(1).ToString();
            this.StartDatum          = dataReader.GetValue(2).ToString();
            this.EndDatum            = dataReader.GetValue(3).ToString();
            this.AnsprechpartnerID   = dataReader.GetValue(4).ToString();
            this.AnzahlStunden       = dataReader.GetValue(5).ToString();
            this.Verrechnungssatz    = dataReader.GetValue(6).ToString();
            this.ProjektTitel        = dataReader.GetValue(7).ToString();
            this.Gesprächsperson     = dataReader.GetValue(8).ToString();
            this.Disponent           = dataReader.GetValue(9).ToString();
            this.ProjektBeschreibung = dataReader.GetValue(10).ToString();
        }
    }
}
