using ContractApplikation.Src.Helper;
using ContractApplikation.Src.Model;
using System.Collections.Generic;

namespace ContractApplikation.Src.Controller
{
    public class DataManager
    {
        public List<Ansprechpartner> CustomerList { get; private set; }
        public List<Projekt> ProjectList { get; private set; }

        public DataManager()
        {
            CustomerList = OleDbHelper.FetchCustomerDetails();
            ProjectList  = OleDbHelper.FetchProjectDetails();
        }

        public void AddCustomer(Ansprechpartner customer)
        {
            CustomerList.Add(customer);
            OleDbHelper.InsertCustomerDetail(customer);
        }

        public void AddProject(Projekt project)
        {
            ProjectList.Add(project);
            OleDbHelper.InsertProjectDetail(project);
        }

        public Ansprechpartner CustomerForIndex(int index)
        {
            return CustomerList[index];
        }

        public Projekt ProjektForIndex(int index)
        {
            return ProjectList[index];
        }
    }
}
