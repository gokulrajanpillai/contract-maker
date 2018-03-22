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

        public bool AddCustomer(Ansprechpartner customer)
        {
            CustomerList.Add(customer);
            return OleDbHelper.InsertCustomerDetail(customer);
        }

        public bool AddProject(Projekt project)
        {
            ProjectList.Add(project);
            return OleDbHelper.InsertProjectDetail(project);
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
