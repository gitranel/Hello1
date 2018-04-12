using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.Diagnostics;
using System.Net;
using Aspose.Email.Mail;
using Aspose.Email.Exchange;

namespace AsposeMail1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            //label1.Text = "";
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // ExStart:SendEmailMessagesUsingExchangeServer
            // Create instance of ExchangeClient class by giving credentials
            //ExchangeClient client = new ExchangeClient("https://MachineName/exchange/username", "username", "password", "domain");
            ExchangeClient client = new ExchangeClient("https://outlook.ncs.corp.int-ads", "ranelaa", "_pa55111", "ncs");

            // Create instance of type MailMessage
            MailMessage msg = new MailMessage();
            //msg.Sender = "ranelaa@ncs.com.sg";
            msg.From = "ranelaa@ncs.com.sg";
            msg.To = "ranelaa@ncs.com.sg";
            msg.Subject = "Hrconnect2 Mail -Aspose";
            msg.HtmlBody = "<h3>sending message from exchange server</h3>";

            try
            {
                // Send the message
                client.Send(msg);
            }
            catch (Exception ex)
            {
                Console.Write( ex.Message);
            }
             //ExEnd:SendEmailMessagesUsingExchangeServer
            //label1.Text = "Sent";


            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // From
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            // To
        }

        private void label1_Click(object sender, EventArgs e)
        {
            //
        }

        private void button2_Click(object sender, EventArgs e)
        {

            //https://outlook.ncs.corp.int-ads/ews/Services.wsdl
            // ExStart:SendEmailMessagesUsingExchangeServer
            // Create instance of IEWSClient class by giving credentials

            //IEWSClient client = EWSClient.GetEWSClient("https://outlook.ncs.corp.int-ads/ews/Services.wsdl", "ranelaa", "_pa55111", "ncs");

            const string mailboxUri = "https://outlook.ncs.corp.int-ads/ews/Services.wsdl";
            //const string mailboxUri = "https://outlook.ncs.corp.int-ads";
            const string domain = @"ncs";
            const string username = @"ranelaa";
            const string password = @"_pa55111";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);


            // Create instance of type MailMessage
            MailMessage msg = new MailMessage();
            msg.Sender = "ranelaa@ncs.com.sg";
            msg.From = "ranelaa@ncs.com.sg";
            msg.To = "ranelaa@ncs.com.sg";
            msg.Subject = "Hrconnect2 Mail -Aspose EWS";
            msg.HtmlBody = "<h3>sending message from exchange server</h3>";




            // Send the message
            try
            {
                // Send the message
                client.Send(msg);
            }
            catch (Exception ex)
            {
                //throw ex;
                Console.Write(ex.Message);
            }
             //ExEnd:SendEmailMessagesUsingExchangeServer
        }

        private void button3_Click(object sender, EventArgs e)
        {

            // Create an instance of SmtpClient class
            SmtpClient client = new SmtpClient();

            // Specify your mailing host server, Username, Password, Port # and Security option
            client.Host = "smtpdev1.ncs.com.sg";
            client.Username = "username";
            client.Password = "password";
            client.Port = 587;
            //client.SecurityOptions = SecurityOptions.SSLExplicit;
            try
            {


                // Create instance of type MailMessage
                MailMessage msg = new MailMessage();
                msg.Sender = "ranelaa@ncs.com.sg";
                msg.From = "ranelaa@ncs.com.sg";
                msg.To = "ranelaa@ncs.com.sg";
                msg.Subject = "Hrconnect2 Mail -Aspose SMTP";
                msg.HtmlBody = "<h3>sending message from exchange server</h3>";

                // Client.Send will send this message
                client.Send(msg);
                Console.WriteLine("Message sent");
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
            Console.WriteLine("Press enter to quit");
            Console.Read();
        }
    }
}