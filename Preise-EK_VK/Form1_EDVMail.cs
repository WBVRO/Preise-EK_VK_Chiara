using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.IO;
using System.Collections;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;


//20210920 - Abfrage ob E-Mail erstellt wird. //Zeilen = 1 - dann nicht = Überschrift
//20221122 - Nur Chiara als Empfänger


namespace Preise_EK_VK
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Connection öffenen
            var conn = new OdbcConnection();
            conn.ConnectionString =
                "DSN=Riedgruppe_32;" +
                "UID=SYSDBA;" +
                "PWD=masterkey;";
            conn.Open();

            
            //Einkaufspreis GDI in Bestellung neu hinterlegt - prüfen
            ArrayList myArrayList = new ArrayList();
            myArrayList.Add("SELECT T2.ARTIKELNR, ARTIKEL.Suchname, Max(T2.MATERIAL) As VK_Preis, T3.EK_Preis "
            + "From PREISE As T2 "
            + "Inner Join ( SELECT ARTIKELNR, Max(MATERIAL) As EK_Preis "
            + "From PREISE "
            + "WHERE PREISLST = 0  "
            + "GROUP BY ARTIKELNR ) As T3 "
            + "ON T2.ARTIKELNR = T3.ARTIKELNR "
            + "INNER JOIN Artikel "
            + "ON T2.ARTIKELNR = ARTIKEL.ARTIKELNR "
            + "Where T2.PREISLST = 1 "
            + "GROUP BY T2.ARTIKELNR, T3.EK_Preis, ARTIKEL.Suchname "
            + "HAVING Max(T2.MATERIAL) < T3.EK_Preis;");
            ////String Objekt - 
            object[] myStringArray = myArrayList.ToArray();
            string strSQL = myStringArray[0].ToString();
            for (int i = 1; i <= myStringArray.Length; i++)
            {
                //textBox2.Text = myStringArray[i - 1].ToString();
            }

            OdbcCommand cmd = new OdbcCommand(strSQL, conn);
            OdbcDataReader readerTest = cmd.ExecuteReader();
            int Count = 0;
            while (readerTest.Read())
            {
                Count += 1;
            }
            readerTest.Close();

            OdbcDataReader reader = cmd.ExecuteReader(); //1.Reader für Grid wird in Table geladen
            DataTable dataTable = new DataTable();
            dataTable.Load(reader);
            dataGridView1.DataSource = dataTable;
            //conn.Close();
            //

            //Datum für SQL Abfrage
            //DateTime dt = new DateTime();
            DateTime dt = DateTime.Now;
            dt = dt.AddDays(-5);
            String sDatum =dt.ToString();
            

            //Einkaufspreis GDI in Rechnung prüfen
            ArrayList myArrayList2 = new ArrayList();
            myArrayList2.Add("SELECT T2.ARTIKELNR, ARTIKEL.Suchname, Max(T2.MATERIAL) As VK_Preis, T3.EK_Preis "
            + "From PREISE As T2 "
            + "Inner Join (SELECT ARTIKELNR, Max(MATERIAL) As EK_Preis "
            + "From PREISE "
            + "WHERE PREISLST = 0  "
            + "GROUP BY ARTIKELNR ) As T3 "
            + "ON T2.ARTIKELNR = T3.ARTIKELNR "            
            + "INNER JOIN (SELECT ARTIKELNR From BELEGPOS WHERE BELEGTYP = 'E' AND BELEGART = 'RE' "
            + "AND  CREATEDATUM >= '" + sDatum + "') AS BPOS " 
            + "On T2.ARTIKELNR = BPOS.ARTIKELNR "
            + "INNER JOIN Artikel "
            + "ON T2.ARTIKELNR = ARTIKEL.ARTIKELNR "
            + "Where T2.PREISLST = 1 "
            + "GROUP BY T2.ARTIKELNR, T3.EK_Preis, ARTIKEL.Suchname "
            + "HAVING Max(T2.MATERIAL) < T3.EK_Preis;");

            ////String Objekt - 
            object[] myStringArray2 = myArrayList2.ToArray();
            string strSQL2 = myStringArray2[0].ToString();
            for (int i = 1; i <= myStringArray2.Length; i++)
            {
                //textBox2.Text = myStringArray[i - 1].ToString();
            }

            OdbcCommand cmd2 = new OdbcCommand(strSQL2, conn);
            OdbcDataReader readerTest2 = cmd2.ExecuteReader();
            int Count2 = 0;
            while (readerTest2.Read())
            {
                Count2 += 1;
            }
            readerTest2.Close();

            OdbcDataReader reader2 = cmd2.ExecuteReader(); //2.Reader für Grid wird in Table geladen
            DataTable dataTable2 = new DataTable();
            dataTable2.Load(reader2);
            dataGridView2.DataSource = dataTable2;


            try
            {
                //Email versenden
                MailMessage Email = new MailMessage();
                //lager@riedgruppe-ost.de
                MailAddress Sender = new MailAddress("lager@riedgruppe-ost.de");
                //Absender einstellen 
                Email.From = Sender;
                //Betreff hinzufügen
                Email.Subject = "GDI - VK EK Preis anpassen";
                //HTML Ansicht
                Email.IsBodyHtml = true;
                //MailBody
                string mailBody = "";                

                // Empfänger hinzufügen - EDV immer wegen Check ob Mail kommt
                // Email.To.Add("edv@riedgruppe-ost.de");

                if (dataGridView2.RowCount > 1)
                {
                    //Empfänger hinzufügen - Wer noch außer EDV? 
                    //Email.To.Add("technik@riedgruppe-ost.de");
                    Email.To.Add("chiara.termini@riedgruppe-ost.de");

                    //Nachrichtentext hinzufügen  
                    mailBody = "Preis anpassen für in RG übernomme Artikel:";
                    //string mailBody = "";
                    //mailBody += @"\\wbvro-lager01\GDI\Mindestbestand_CSV";
                    mailBody += "<table width='100%' style='border:Solid 1px Black;'>";
                    // Column headers
                    string columnsHeader = "";
                    for (int i = 0; i < dataGridView2.Columns.Count; i++)
                    {
                        columnsHeader += "<td>" + dataGridView2.Columns[i].Name + "</td>";
                    }
                    mailBody += columnsHeader + "</td>";

                    // Zeilen:
                    foreach (DataGridViewRow row in dataGridView2.Rows)
                    {
                        mailBody += "<tr>";
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            mailBody += "<td>" + cell.Value + "</td>";
                        }
                        mailBody += "</tr>";
                    }

                    mailBody += "</table>";                    
                    //your rest of the original code
                    Email.Body = mailBody;

                }


                //20210920 - Abfrage ob E-Mail erstellt wird. //Zeilen = 1 - dann nicht = Überschrift
                else if (dataGridView1.RowCount > 1)
                {
                    //Ab Buchung Einkaufspreis GDI in Bestellung 

                    //Empfänger hinzufügen - Wer noch außer EDV
                    //Email.To.Add("technik@riedgruppe-ost.de");
                    Email.To.Add("chiara.termini@riedgruppe-ost.de");

                    //Nachrichtentext hinzufügen 
                    mailBody += "Alte Version zur Kontrolle - kann irgendwann raus, wenn das andere korrekt ist -";
                    //Nachrichtentext hinzufügen  
                    mailBody += "<table width='100%' style='border:Solid 1px Black;'>";
                    // Column headers
                    string columnsHeader = "";
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        columnsHeader += "<td>" + dataGridView1.Columns[i].Name + "</td>";
                    }
                    mailBody += columnsHeader + "</td>";

                    // Zeilen:
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        mailBody += "<tr>";
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            mailBody += "<td>" + cell.Value + "</td>";
                        }
                        mailBody += "</tr>";
                    }

                    mailBody += "</table>";
                    //your rest of the original code
                    Email.Body = mailBody;
                }

                else
                {
                    //Check ob Task ausgeführt für EDV
                    //Empfänger hinzufügen
                    //EDV ist oben drin weil die Mail immer kommen soll
                    //Nachrichtentext hinzufügen  
                    Email.Body = "Aufgabe ausgeführt";
                }

                //Mail senden
                SmtpClient MailClient = new SmtpClient("mail.riedgruppe-ost.de", 25); // Postausgangsserver definieren
                MailClient.Credentials = new NetworkCredential("edv@riedgruppe-ost.de", "Server2014!");   //Credentials; // Anmeldeinformationen setzen
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                MailClient.Send(Email); // Email senden

            }


            catch (SmtpException exeption)
            {
                MessageBox.Show(exeption.Message);
            }



            Application.Exit();
        }
    }
}
