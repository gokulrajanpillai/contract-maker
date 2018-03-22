using ContractApplikation.Src.Helper;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Reflection;
using System.Windows.Forms;

namespace ContractApplikation.Src.Model
{
    public class Projekt
    {
        public String Projektnummer { get; private set; }

        public String StartDatum { get; private set; }

        public String EndDatum { get; private set; }

        public Int32 AnsprechpartnerID { get; private set; }

        public Int32 AnzahlStunden { get; private set; }

        public Int32 Verrechnungssatz { get; private set; }

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
                decimal payment = Verrechnungssatz;
                return Utilities.AddCurrencySymbol(decimal.Round(payment, 2, MidpointRounding.AwayFromZero).ToString());
            }
        }

        // Custom property: not part of the database
        public String AngebotSumme
        {
            get
            {
                decimal hours = AnzahlStunden;
                decimal payment = Verrechnungssatz;
                decimal sum = hours * payment;
                return Utilities.AddCurrencySymbol(decimal.Round(sum, 2, MidpointRounding.AwayFromZero).ToString());
            }
        }

        public Projekt(List<TextBox> listOfTextboxes)
        {
            foreach (TextBox textBox in listOfTextboxes)
            {
                PropertyInfo property = this.GetType().GetProperty(Utilities.FirstLetterToUpperCase(textBox.Name));

                if (property.PropertyType == typeof(Int32))
                {
                    property.SetValue(this, int.Parse(textBox.Text));
                }
                else
                {
                    property.SetValue(this, textBox.Text);
                }
            }
        }

        public Projekt(OleDbDataReader dataReader)
        {
            this.Projektnummer          = dataReader.GetValue(1).ToString();
            this.StartDatum             = dataReader.GetValue(2).ToString();
            this.EndDatum               = dataReader.GetValue(3).ToString();
            this.AnsprechpartnerID      = dataReader.GetInt32(4);
            this.AnzahlStunden          = dataReader.GetInt32(5);
            this.Verrechnungssatz       = dataReader.GetInt32(6);
            this.Koordinator            = dataReader.GetValue(11).ToString();
            this.Gesprächsperson        = dataReader.GetValue(8).ToString();
            this.Disponent              = dataReader.GetValue(9).ToString();
            this.ProjektTitel           = dataReader.GetValue(7).ToString();
            this.ProjektBeschreibung    = dataReader.GetValue(10).ToString();
        }
    }
}
