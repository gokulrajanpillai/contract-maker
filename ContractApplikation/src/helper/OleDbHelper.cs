using ContractApplikation.Src.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ContractApplikation.Src.Helper
{
    public class OleDbHelper
    {
        private static readonly string CONNECTION_STRING = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="+Constants.FileLocation.DATASOURCE;

        private static OleDbConnection OpenConnection()
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = CONNECTION_STRING;
            conn.Open();
            return conn;
        }

        private static void CloseConnection(OleDbConnection conn)
        {
            conn.Close();
        }

        private static void AddCustomerDetailToDatabase(Ansprechpartner kunden, OleDbConnection conn)
        {
            var cmd = new OleDbCommand("INSERT INTO Ansprechpartner " +
                    "(Anrede, Vorname, Nachname, Abteilung, Email, Telefon, Strasse, PLZ, Ort, Firma, Abteilungszusatz, Geschäftsbereich) " +
                    "VALUES (@Anrede, @Vorname, @Nachname, @Abteilung, @Email, @Telefon, @Strasse, @PLZ, @Ort, @Firma, @Abteilungszusatz, @Geschäftsbereich)");
            cmd.Connection = conn;

            if (cmd.Connection.State == System.Data.ConnectionState.Open)
            {
                cmd.Parameters.Add("@Anrede", OleDbType.VarChar).Value              = kunden.Anrede;
                cmd.Parameters.Add("@Vorname", OleDbType.VarChar).Value             = kunden.Vorname;
                cmd.Parameters.Add("@Nachname", OleDbType.VarChar).Value            = kunden.Nachname;
                cmd.Parameters.Add("@Abteilung", OleDbType.VarChar).Value           = kunden.Abteilung;
                cmd.Parameters.Add("@Email", OleDbType.VarChar).Value               = kunden.Email;
                cmd.Parameters.Add("@Telefon", OleDbType.VarChar).Value             = kunden.Telefon;
                cmd.Parameters.Add("@Strasse", OleDbType.VarChar).Value             = kunden.Strasse;
                cmd.Parameters.Add("@PLZ", OleDbType.VarChar).Value                 = kunden.PLZ;
                cmd.Parameters.Add("@Ort", OleDbType.VarChar).Value                 = kunden.Ort;
                cmd.Parameters.Add("@Firma", OleDbType.VarChar).Value               = kunden.Firma;
                cmd.Parameters.Add("@Abteilungszusatz", OleDbType.VarChar).Value    = kunden.Abteilungszusatz;
                cmd.Parameters.Add("@Geschäftsbereich", OleDbType.VarChar).Value    = kunden.Geschäftsbereich;
                cmd.ExecuteNonQuery();
                MessageBox.Show("Customer details of " + kunden.Vorname + " is successfully entered to the database.");
            }
        }

        private static void AddProjectDetailToDatabase(Projekt project, OleDbConnection conn)
        {
            var cmd = new OleDbCommand("INSERT INTO Projekt " +
                    "(Projektnummer, StartDatum, EndDatum, AnsprechpartnerID, AnzahlStunden, Verrechnungssatz, Koordinator, Gesprächsperson, Disponent, ProjektTitel, ProjektBeschreibung) " +
                    "VALUES (@Projektnummer, @StartDatum, @EndDatum, @AnsprechpartnerID, @AnzahlStunden, @Verrechnungssatz, @Koordinator, @Gesprächsperson, @Disponent, @ProjektTitel, @ProjektBeschreibung)");
            cmd.Connection = conn;

            if (cmd.Connection.State == System.Data.ConnectionState.Open)
            {
                cmd.Parameters.Add("@Projektnummer", OleDbType.VarChar).Value       = project.Projektnummer;
                cmd.Parameters.Add("@StartDatum", OleDbType.VarChar).Value          = project.StartDatum;
                cmd.Parameters.Add("@EndDatum", OleDbType.VarChar).Value            = project.EndDatum;
                cmd.Parameters.Add("@AnsprechpartnerID", OleDbType.Integer).Value   = project.AnsprechpartnerID;
                cmd.Parameters.Add("@AnzahlStunden", OleDbType.VarChar).Value       = Utilities.RoundByTwoDecimalPlaces(project.AnzahlStunden).ToString();
                cmd.Parameters.Add("@Verrechnungssatz", OleDbType.VarChar).Value    = Utilities.RoundByTwoDecimalPlaces(project.Verrechnungssatz).ToString();
                cmd.Parameters.Add("@Koordinator", OleDbType.VarChar).Value         = project.Koordinator;
                cmd.Parameters.Add("@Gesprächsperson", OleDbType.VarChar).Value     = project.Gesprächsperson;
                cmd.Parameters.Add("@Disponent", OleDbType.VarChar).Value           = project.Disponent;
                cmd.Parameters.Add("@ProjektTitel", OleDbType.VarChar).Value        = project.ProjektTitel;
                cmd.Parameters.Add("@ProjektBeschreibung", OleDbType.VarChar).Value = project.ProjektBeschreibung;
                cmd.ExecuteNonQuery();
                MessageBox.Show("Project details of " + project.ProjektTitel + " is successfully entered to the database.");
            }
        }

        public static bool InsertCustomerDetail(Ansprechpartner kunden)
        {
            try
            {
                OleDbConnection conn = OpenConnection();
                AddCustomerDetailToDatabase(kunden, conn);
                CloseConnection(conn);
            }
            catch(Exception e)
            {
                MessageBox.Show("Error: " + e.Message);
                return false;
            }

            return true;
        }

        public static bool InsertProjectDetail(Projekt projekt)
        {
            try
            {
                OleDbConnection conn = OpenConnection();
                AddProjectDetailToDatabase(projekt, conn);
                CloseConnection(conn);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message);
                return false;
            }

            return true;
        }

        public static List<Ansprechpartner> FetchCustomerDetails()
        {
            List<Ansprechpartner> customerList = new List<Ansprechpartner>();

            OleDbConnection oleDbConnection = OpenConnection();
            OleDbDataAdapter oledbAdapter;
            DataSet dataSet = new DataSet();
            DataTable dataTable;

            try
            {
                oledbAdapter = new OleDbDataAdapter("SELECT * FROM Ansprechpartner", oleDbConnection);
                oledbAdapter.Fill(dataSet, "Ansprechpartner");
                oledbAdapter.Dispose();

                dataTable = dataSet.Tables[0];
                foreach (DataRow row in dataTable.Rows)
                {
                    customerList.Add(new Ansprechpartner(row));
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Can not open connection ! "+e.Message);
            }
            finally
            {
                CloseConnection(oleDbConnection);
            }

            return customerList;
        }

        
        public static List<Projekt> FetchProjectDetails()
        {
            List<Projekt> projektList = new List<Projekt>();

            OleDbConnection oleDbConnection = OpenConnection();
            OleDbDataAdapter oledbAdapter;
            DataSet dataSet = new DataSet();
            DataTable dataTable;

            try
            {
                oledbAdapter = new OleDbDataAdapter("SELECT * FROM Projekt", oleDbConnection);
                oledbAdapter.Fill(dataSet, "Projekt");
                oledbAdapter.Dispose();

                dataTable = dataSet.Tables[0];
                foreach (DataRow row in dataTable.Rows)
                {
                    projektList.Add(new Projekt(row));
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Can not open connection ! "+e.Message);
            }
            finally
            {
                CloseConnection(oleDbConnection);
            }

            return projektList;
        }
    }
}
