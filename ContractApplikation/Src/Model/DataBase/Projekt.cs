using ContractApplikation.Src.Helper;
using System;
using System.Collections.Generic;
using System.Data;
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

        public Decimal AnzahlStunden { get; private set; }

        public Decimal Verrechnungssatz { get; private set; }

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
                return Utilities.AddCurrencySymbol(Utilities.RoundByTwoDecimalPlaces(payment).ToString());
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
                return Utilities.AddCurrencySymbol(Utilities.RoundByTwoDecimalPlaces(sum).ToString());
            }
        }

        // Custom property: not part of the database
        public String CostTableFileName
        {
            get
            {
                return "kostTabelle_" + ProjektTitel + ".xlsx";
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
                else if (property.PropertyType == typeof(decimal))
                {
                    property.SetValue(this, decimal.Parse(textBox.Text));
                }
                else
                {
                    property.SetValue(this, textBox.Text);
                }
            }
        }

        public Projekt(DataRow dataRow)
        {
            this.Projektnummer = dataRow["Projektnummer"].ToString();
            this.StartDatum = dataRow["StartDatum"].ToString();
            this.EndDatum = dataRow["EndDatum"].ToString();
            this.AnsprechpartnerID = Int32.Parse(dataRow["AnsprechpartnerID"].ToString());
            this.AnzahlStunden = Decimal.Parse(dataRow["AnzahlStunden"].ToString());
            this.Verrechnungssatz = Decimal.Parse(dataRow["Verrechnungssatz"].ToString());
            this.Koordinator = dataRow["Koordinator"].ToString();
            this.Gesprächsperson = dataRow["Gesprächsperson"].ToString();
            this.Disponent = dataRow["Disponent"].ToString();
            this.ProjektTitel = dataRow["ProjektTitel"].ToString();
            this.ProjektBeschreibung = dataRow["ProjektBeschreibung"].ToString();
        }
    }
}
