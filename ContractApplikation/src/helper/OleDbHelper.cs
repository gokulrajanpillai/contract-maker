using ContractApplikation.Src.Model;
using System;
using System.Collections.Generic;
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
                    "(Projektnummer, StartDatum, EndDatum, AnsprechpartnerID, AnzahlStunden, Verrechnungssatz, ProjektTitel, Gesprächsperson, Disponent, ProjektBeschreibung) " +
                    "VALUES (@Projektnummer, @StartDatum, @EndDatum, @AnsprechpartnerID, @AnzahlStunden, @Verrechnungssatz, @ProjektTitel, @Gesprächsperson, @Disponent, @ProjektBeschreibung)");
            cmd.Connection = conn;

            if (cmd.Connection.State == System.Data.ConnectionState.Open)
            {
                cmd.Parameters.Add("@Projektnummer", OleDbType.VarChar).Value       = project.Projektnummer;
                cmd.Parameters.Add("@StartDatum", OleDbType.VarChar).Value          = project.StartDatum;
                cmd.Parameters.Add("@EndDatum", OleDbType.VarChar).Value            = project.EndDatum;
                cmd.Parameters.Add("@AnsprechpartnerID", OleDbType.Integer).Value   = project.AnsprechpartnerID;
                cmd.Parameters.Add("@AnzahlStunden", OleDbType.Integer).Value       = project.AnzahlStunden;
                cmd.Parameters.Add("@Verrechnungssatz", OleDbType.Integer).Value    = project.Verrechnungssatz;
                cmd.Parameters.Add("@ProjektTitel", OleDbType.VarChar).Value        = project.ProjektTitel;
                cmd.Parameters.Add("@Gesprächsperson", OleDbType.VarChar).Value     = project.Gesprächsperson;
                cmd.Parameters.Add("@Disponent", OleDbType.VarChar).Value           = project.Disponent;
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
            try
            {
                List<Ansprechpartner> customerList = new List<Ansprechpartner>();
                OleDbDataReader dr = null;
                var cmd = new OleDbCommand("SELECT * FROM Ansprechpartner", OpenConnection());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    customerList.Add(new Ansprechpartner(dr));
                }

                return customerList;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message);
            }

            return null;
        }

        public static List<Projekt> FetchProjectDetails()
        {
            try
            {
                List<Projekt> projectList = new List<Projekt>();
                OleDbDataReader dr = null;
                var cmd = new OleDbCommand("SELECT * FROM Projekt", OpenConnection());
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    projectList.Add(new Projekt(dr));
                }

                return projectList;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message);
            }

            return null;
        }
    }
}
