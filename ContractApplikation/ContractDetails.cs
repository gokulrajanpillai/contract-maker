using System;
using System.Collections.Generic;
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

        #region Form Events
        
        #region Init
        public ContractDetails()
        {
            InitializeComponent();
        }
        #endregion

        #region Load Event
        private void ContractDetails_Load(object sender, EventArgs e)
        {
            backgrdDBWorker.RunWorkerAsync();
        }
        #endregion

        #region Form Refresh
        private void RefreshForm()
        {
            Utilities.ClearControls(this.Controls);
        }
        #endregion

        #endregion

        #region Information Generation Helper Function
        private List<System.Windows.Forms.TextBox> ListOfTextBoxFromControlCollection(Control.ControlCollection controlsForPage)
        {
            IEnumerable<System.Windows.Forms.TextBox> textboxControls = controlsForPage.OfType<System.Windows.Forms.TextBox>();
            return textboxControls.ToList();
        }
        #endregion

        #region Customer Information Generation

        private Ansprechpartner GenerateCustomerWithControl(Control.ControlCollection controlsForCustomerTabPage)
        {
            return new Ansprechpartner(ListOfTextBoxFromControlCollection(controlsForCustomerTabPage), GetSalutationForCustomer());
        }

        private Salutation GetSalutationForCustomer()
        {
            return (herrRadioBtn.Checked ? Salutation.HERR : Salutation.FRAU);
        }

        private bool CustomerDetailIsValid()
        {
            var controlsForCustomerTabPage = this.Controls[0].Controls[0].Controls;

            System.Windows.Forms.TextBox emptyItem = controlsForCustomerTabPage.OfType<System.Windows.Forms.TextBox>().FirstOrDefault(tb => String.IsNullOrWhiteSpace(tb.Text));
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

        #endregion

        #region Project Information Generation

        private string RemoveTimeFromDateString(string dateString)
        {
            string finalString = dateString;
            if (dateString.Contains(" "))
                finalString = dateString.Substring(0, dateString.IndexOf(' '));

            return finalString;
        }

        private Projekt GenerateProjectWithControl(Control.ControlCollection controlsForProjectTabPage)
        {
            List<System.Windows.Forms.TextBox> textboxes = ListOfTextBoxFromControlCollection(controlsForProjectTabPage);
            textboxes.Add(Utilities.GenerateTextBoxWithNameAndValue("startDatum", RemoveTimeFromDateString(startDatumDtPikr.Value.ToString())));
            textboxes.Add(Utilities.GenerateTextBoxWithNameAndValue("endDatum", RemoveTimeFromDateString(endDatumDtPikr.Value.ToString())));
            textboxes.Add(Utilities.GenerateTextBoxWithNameAndValue("ansprechpartnerID", ansprechpartnerComboBox.SelectedIndex.ToString()));
            return new Projekt(textboxes);
        }

        private bool ProjectDetailIsValid()
        {
            var controlsForProjectTabPage = this.Controls[0].Controls[1].Controls;

            System.Windows.Forms.TextBox emptyItem = controlsForProjectTabPage.OfType<System.Windows.Forms.TextBox>().FirstOrDefault(tb => String.IsNullOrWhiteSpace(tb.Text));
            if (emptyItem != null)
            {
                MessageBox.Show("Geben Sie den " + emptyItem.Name + " ein");
                return false;
            }
            else
            {
                return AreProjectDatesValid();
            }
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

        #endregion

        #region Background Worker

        #region Do Work
        private void BackgrdDBWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            model = new DataManager();
        }
        #endregion

        #region Completion
        private void BackgrdDBWorker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            UpdateComboBoxValues();
        }
        #endregion

        #region Update ComboBoxes

        private void UpdateComboBoxValues()
        {
            UpdateCustomerComboBox();
            UpdateProjectComboBox();
        }

        private void UpdateProjectComboBox()
        {
            projektComboBox.Items.Clear();
            kost_projectComboBox.Items.Clear();

            foreach (Projekt proj in model.ProjectList)
            {
                projektComboBox.Items.Add(new CustomComboBoxItem(proj.ProjektTitel, proj));
                kost_projectComboBox.Items.Add(new CustomComboBoxItem(proj.ProjektTitel, proj));
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

        #endregion

        #endregion

        #region UI Interactions

        #region Project Combobox Selection
        private void ProjektComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Projekt proj        = model.ProjektForIndex(projektComboBox.SelectedIndex);
            contractName.Text   = proj.ProjektTitel;
        }
        #endregion

        #region TabPage Selection
        private void TabPage_Selected(object sender, TabControlEventArgs e)
        {
            if ((sender as TabControl).SelectedTab.Name.Equals("ProjektkostenTabelle"))
            {
                
            }
        }

        
        #endregion

        #region Button Click Events

        private void CreateCustomerButtonClicked(object sender, EventArgs e)
        {
            if (CustomerDetailIsValid())
            {
                var controlsForCustomerTabPage = this.Controls[0].Controls[0].Controls;
                if (model.AddCustomer(GenerateCustomerWithControl(controlsForCustomerTabPage)))
                {
                    UpdateCustomerComboBox();
                    RefreshForm();
                }
            }
        }


        private void CreateProjectButtonClicked(object sender, EventArgs e)
        {
            if (ProjectDetailIsValid())
            {
                var controlsForProjectTabPage = this.Controls[0].Controls[1].Controls;
                if (model.AddProject(GenerateProjectWithControl(controlsForProjectTabPage)))
                {
                    UpdateProjectComboBox();
                    RefreshForm();
                }
            }
        }


        private void CreateContractButtonClicked(object sender, EventArgs e)
        {
            Projekt proj = model.ProjektForIndex(projektComboBox.SelectedIndex);
            Ansprechpartner kunden = model.CustomerForIndex(proj.AnsprechpartnerID);
            DocumentManager.GenerateContractDocument(contractName.Text + ".docx", kunden, proj);
        }


        private void EditProjectCostTable_Click(object sender, EventArgs e)
        {
            DocumentManager.EditCostTableForProject(model.ProjektForIndex(kost_projectComboBox.SelectedIndex));
            // Class1.WordFromExcel(Constants.FileLocation.PROTOTYPE_COSTTABLE);
        }

        #endregion

        #endregion

    }
}
