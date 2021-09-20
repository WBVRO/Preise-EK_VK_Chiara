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
            var conn = new OdbcConnection();
            conn.ConnectionString =
                "DSN=Riedgruppe_32;" +
                "UID=SYSDBA;" +
                "PWD=masterkey;";
            conn.Open();

            ArrayList myArrayList = new ArrayList();
            myArrayList.Add("SELECT T2.ARTIKELNR, Max(T2.MATERIAL) As VK_Preis, T3.EK_Preis "
            + "From PREISE As T2 "
            + "Inner Join ( SELECT ARTIKELNR, Max(MATERIAL) As EK_Preis "
            + "From PREISE "
            + "WHERE PREISLST = 0  "                           
            + "GROUP BY ARTIKELNR ) As T3 "                            
            + "ON T2.ARTIKELNR = T3.ARTIKELNR "
            + "Where T2.PREISLST = 1 "                                                         
            + "GROUP BY T2.ARTIKELNR, T3.EK_Preis "
            + "HAVING Max(T2.MATERIAL) < EK_Preis");
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

            try
            {
                //Email versenden
                MailMessage Email = new MailMessage();
                //lager@riedgruppe-ost.de
                MailAddress Sender = new MailAddress("lager@riedgruppe-ost.de");
                // Absender einstellen 
                Email.From = Sender;
                // Empfänger hinzufügen
                Email.To.Add("lager@riedgruppe-ost.de");
                // Betreff hinzufügen
                Email.Subject = "VK Preis anpassen";
                // HTML Ansicht
                Email.IsBodyHtml = true;
                // Nachrichtentext hinzufügen  
                string mailBody = "";
                //mailBody += @"\\wbvro-lager01\GDI\Mindestbestand_CSV";
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
                //Anhängen der Datei
                //Attachment attach = new Attachment(@"\\wbvro-lager01\GDI\Mindestbestand_CSV\Mindestbestand.csv");
                SmtpClient MailClient = new SmtpClient("mail.riedgruppe-ost.de", 25); // Postausgangsserver definieren
                //SmtpClient MailClient = new SmtpClient("192.168.0.93");
                MailClient.Credentials = new NetworkCredential("wbv@riedgruppe-ost.de", "Server2014!");   //Credentials; // Anmeldeinformationen setzen*/
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                MailClient.Send(Email); // Email senden             
            }
            catch (SmtpException exeption)
            {
                MessageBox.Show(exeption.Message);
            }

            //Email versenden
            MailMessage Email2 = new MailMessage();
            //lager@riedgruppe-ost.de
            MailAddress Sender2 = new MailAddress("lager@riedgruppe-ost.de");
            // Absender einstellen 
            Email2.From = Sender2;
            // Empfänger hinzufügen
            Email2.To.Add("edv@riedgruppe-ost.de");
            // Betreff hinzufügen
            Email2.Subject = "VK Preis - Aufgabe wurde ausgeführt";
            // HTML Ansicht
            Email2.IsBodyHtml = true;
            // Nachrichtentext hinzufügen  
            string mailBody2 = "Aufgabe wurde ausgeführt";
            Email2.Body = mailBody2;
            SmtpClient MailClient2 = new SmtpClient("mail.riedgruppe-ost.de", 25); // Postausgangsserver definieren
                                                                                  //SmtpClient MailClient = new SmtpClient("192.168.0.93");
            MailClient2.Credentials = new NetworkCredential("wbv@riedgruppe-ost.de", "Server2014!");   //Credentials; // Anmeldeinformationen setzen*/
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            MailClient2.Send(Email2); // Email senden

            Application.Exit();
        }
    }
}
