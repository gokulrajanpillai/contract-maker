using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;
using ContractApplikation.Src.Controller;
using ContractApplikation.Src.Helper;
using ContractApplikation.Src.Model;

namespace ContractApplikation
{

    public partial class ContractDetails : Form
    {
        private DataManager model;

        public ContractDetails()
        {
            InitializeComponent();
        }

        private List<TextBox> ListOfTextBoxFromControlCollection(Control.ControlCollection controlsForPage)
        {
            IEnumerable<TextBox> textboxControls = controlsForPage.OfType<TextBox>();
            return textboxControls.ToList();
        }

        private void CreateCustomerBtnClicked(object sender, EventArgs e)
        {
            if (CustomerDetailIsValid())
            {
                var controlsForCustomerTabPage = this.Controls[0].Controls[0].Controls;
                model.AddCustomer(GenerateCustomerWithControl(controlsForCustomerTabPage));
                UpdateCustomerComboBox();
                RefreshForm();
            }
        }

        private void RefreshForm()
        {
            Utilities.ClearControls(this.Controls);
        }

        private Ansprechpartner GenerateCustomerWithControl(Control.ControlCollection controlsForCustomerTabPage)
        {
            return new Ansprechpartner(ListOfTextBoxFromControlCollection(controlsForCustomerTabPage), GetSalutationForCustomer());
        }

        private Salutation GetSalutationForCustomer()
        {
            return (herrRadioBtn.Checked ? Salutation.HERR : Salutation.FRAU);
        }

        /**
        private void AddCustomDetailToDatabase(Ansprechpartner kunden)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\GRajan\source\repos\WindowsFormsApp1\WindowsFormsApp1\Vertrag-DB.accdb";
            conn.Open();
            var cmd = new OleDbCommand("INSERT INTO Ansprechpartner " +
                    "(Bezeichnung, Vorname, Nachname, Abteilung, Email, Telefon, Strasse, Ort, Firma, Abteilungszusatz, Geschäftsbereich) " +
                    "VALUES (@Bezeichnung, @Vorname, @Nachname, @Abteilung, @Email, @Telefon, @Strasse, @Ort, @Firma, @Abteilungszusatz, @Geschäftsbereich)");
            cmd.Connection = conn;

            if (conn.State == System.Data.ConnectionState.Open)
            {
                cmd.Parameters.Add("@Bezeichnung", OleDbType.VarChar).Value      = kunden.bezeichnung;
                cmd.Parameters.Add("@Vorname", OleDbType.VarChar).Value          = kunden.vorname;
                cmd.Parameters.Add("@Nachname", OleDbType.VarChar).Value         = kunden.nachname;
                cmd.Parameters.Add("@Abteilung", OleDbType.VarChar).Value        = kunden.abteilung;
                cmd.Parameters.Add("@Email", OleDbType.VarChar).Value            = kunden.email;
                cmd.Parameters.Add("@Telefon", OleDbType.VarChar).Value          = kunden.telefon;
                cmd.Parameters.Add("@Strasse", OleDbType.VarChar).Value          = kunden.strasse;
                cmd.Parameters.Add("@Ort", OleDbType.VarChar).Value              = kunden.ort;
                cmd.Parameters.Add("@Firma", OleDbType.VarChar).Value            = kunden.firma;
                cmd.Parameters.Add("@Abteilungszusatz", OleDbType.VarChar).Value = kunden.abteilungszusatz;
                cmd.Parameters.Add("@Geschäftsbereich", OleDbType.VarChar).Value = kunden.geschäftsbereich;

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error: "+e.Message);
                }
                finally
                {
                    conn.Close();
                }
            }
        }

        **/

        private bool CustomerDetailIsValid()
        {
            var controlsForCustomerTabPage = this.Controls[0].Controls[0].Controls;

            TextBox emptyItem = controlsForCustomerTabPage.OfType<TextBox>().FirstOrDefault(tb => String.IsNullOrWhiteSpace(tb.Text));
            if (emptyItem != null)
            {
                MessageBox.Show("Geben Sie den " + emptyItem.Name + " ein");
            }
            else if (!herrRadioBtn.Checked && !frauRadioBtn.Checked)
            {
                MessageBox.Show("Wähle ein Geschlecht aus");
            }
            else
            {
                return true;
            }

            return false;
        }

        private void CreateProjectBtnClicked(object sender, EventArgs e)
        {
            if (ProjectDetailIsValid())
            {
                var controlsForProjectTabPage = this.Controls[0].Controls[1].Controls;
                model.AddProject(GenerateProjectWithControl(controlsForProjectTabPage));
                UpdateProjectComboBox();
            }
        }

        private Projekt GenerateProjectWithControl(Control.ControlCollection controlsForProjectTabPage)
        {
            List<TextBox> textboxes = ListOfTextBoxFromControlCollection(controlsForProjectTabPage);
            textboxes.Add(Utilities.GenerateTextBoxWithNameAndValue("startDatum",   startDatumDtPikr.Value.ToString()));
            textboxes.Add(Utilities.GenerateTextBoxWithNameAndValue("endDatum", endDatumDtPikr.Value.ToString()));
            textboxes.Add(Utilities.GenerateTextBoxWithNameAndValue("ansprechpartnerID", ansprechpartnerComboBox.SelectedIndex.ToString()));
            return new Projekt(textboxes);
        }

        private bool ProjectDetailIsValid()
        {
            var controlsForProjectTabPage = this.Controls[0].Controls[1].Controls;

            TextBox emptyItem = controlsForProjectTabPage.OfType<TextBox>().FirstOrDefault(tb => String.IsNullOrWhiteSpace(tb.Text));
            if (emptyItem != null)
            {
                MessageBox.Show("Geben Sie den " + emptyItem.Name + " ein");
            }
            else if (!AreProjectDatesValid())
            {
                MessageBox.Show("Wähle ein Geschlecht aus");
            }
            else
            {
                return true;
            }

            return false;
        }

        private bool AreProjectDatesValid()
        {
            if (startDatumDtPikr.Text == null || endDatumDtPikr.Text == null)
            {
                return false;
            }
            else if (startDatumDtPikr.Value > endDatumDtPikr.Value)
            {
                MessageBox.Show("Startdatum sollte nach Enddatum sein");
                return false;
            }
            else if (startDatumDtPikr.Value == endDatumDtPikr.Value)
            {
                MessageBox.Show("Startdatum und Enddatum können nicht identisch sein");
                return false;
            }

            return true;
        }

        private void ContractDetails_Load(object sender, EventArgs e)
        {
            backgrdDBWorker.RunWorkerAsync();
        }

        private void BackgrdDBWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            model = new DataManager();
        }

        private void BackgrdDBWorker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            UpdateComboBoxValues();
        }

        /**
        private void BindComboBoxValues()
        {
            projektComboBox.DataSource      = model.ProjectList;
            projektComboBox.DisplayMember   = "ProjektTitel";
            projektComboBox.ValueMember     = null;

            ansprechpartnerComboBox.DataSource      = model.CustomerList;
            ansprechpartnerComboBox.DisplayMember   = "Name";
            ansprechpartnerComboBox.ValueMember     = null;
        }
    **/
    
        private void UpdateComboBoxValues()
        {
            UpdateCustomerComboBox();
            UpdateProjectComboBox();
        }

        private void UpdateProjectComboBox()
        {
            projektComboBox.Items.Clear();
            foreach (Projekt proj in model.ProjectList)
            {
                projektComboBox.Items.Add(new CustomComboBoxItem(proj.ProjektTitel, proj));
            }
        }

        private void UpdateCustomerComboBox()
        {
            ansprechpartnerComboBox.Items.Clear();
            foreach (Ansprechpartner cust in model.CustomerList)
            {
                ansprechpartnerComboBox.Items.Add(new CustomComboBoxItem(cust.Name, model.CustomerList.IndexOf(cust)));
            }
        }

        private void GenerateContractButtonClicked(object sender, EventArgs e)
        {
            Projekt proj            = model.ProjektForIndex(projektComboBox.SelectedIndex);
            Ansprechpartner kunden  = model.CustomerForIndex(Int32.Parse(proj.AnsprechpartnerID));
            DocumentManager.GenerateContractDocument("New.docx", kunden, proj);
        }

    }
}
