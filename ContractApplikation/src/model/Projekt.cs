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

        public String Koordinator { get; private set; }

        public String Gesprächsperson { get; private set; }

        public String Disponent { get; private set; }

        public String ProjektBeschreibung { get; private set; }

        // Custom property: not part of the database
        public String Einzelpreis
        {
            get
            {
                decimal payment = decimal.Parse(Verrechnungssatz);
                return Utilities.AddCurrencySymbol(decimal.Round(payment, 2, MidpointRounding.AwayFromZero).ToString());
            }
        }

        // Custom property: not part of the database
        public String AngebotSumme
        {
            get
            {
                decimal hours = decimal.Parse(AnzahlStunden);
                decimal payment = decimal.Parse(Verrechnungssatz);
                decimal sum = hours * payment;
                return Utilities.AddCurrencySymbol(decimal.Round(sum, 2, MidpointRounding.AwayFromZero).ToString());
            }
        }

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
            this.Koordinator         = dataReader.GetValue(8).ToString();
            this.Gesprächsperson     = dataReader.GetValue(9).ToString();
            this.Disponent           = dataReader.GetValue(10).ToString();
            this.ProjektBeschreibung = dataReader.GetValue(11).ToString();
        }
    }
}
