using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.IO;
using System.Net;
using System.Windows.Forms;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace COMPLETE_FLAT_UI
{
    class emailsend
    {
        public X509CertificateCollection ClientCertificates { get; }

        public string SendEmail(string[] mailTo, string[] mailCC, string[] mailBCC, string subject, string mailBody, string[] attachments, string acnt_no, bool ismailBodyHtml)
        {

            string[] informail = new string[4];
           // informail = getcfmail();

            //string tempFilePath = "";
            List<string> tempFiles = new List<string>();
            // Get setting for SMTP
            string SMTP_SERVER = Common.SMTP_SERVER;
            int SMTP_PORT = Int32.Parse(Common.SMTP_PORT);
            string MAIL_FROM = Common.MAIL_FROM;
            string MAIL_FROM_NAME = Common.MAIL_FROM_NAME;
            string SMTP_USERNAME =Common.SMTP_USERNAME;
            string SMTP_PASSWORD = Common.SMTP_PASSWORD;

            // The subject line of the email
            String SUBJECT = subject;
            // The body of the email
            String BODY = mailBody;

            // Create and build a new MailMessage object

            MailMessage message = new MailMessage();
            string htmlText = BODY;
            message.IsBodyHtml = true;
            message.From = new System.Net.Mail.MailAddress(MAIL_FROM, MAIL_FROM_NAME);
            message.Subject = SUBJECT;
            AlternateView plainView = AlternateView.CreateAlternateViewFromString("Some plaintext", Encoding.UTF8, "text/plain");
            // We have something to show in real old mail clients. 


            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlText, Encoding.UTF8, "text/html");
            message.AlternateViews.Add(plainView);
            message.AlternateViews.Add(htmlView);
            message.Body = htmlText;

            if (attachments != null && attachments.Length > 0)
            {
                foreach (string attachmentsen in attachments)
                {
                    //TODO: Check CC email is valid
                    if (!String.IsNullOrEmpty(attachmentsen))
                    {
                        message.Attachments.Add(new System.Net.Mail.Attachment(attachmentsen));
                    }
                }
            }



            if (mailTo != null && mailTo.Length > 0)
            {
                foreach (string emailTo in mailTo)
                {
                    //TODO: Check CC email is valid
                    if (!String.IsNullOrEmpty(emailTo))
                    {
                        message.To.Add(emailTo);
                    }
                }
            }

            if (mailCC != null && mailCC.Length > 0)
            {
                foreach (string emailCc in mailCC)
                {
                    //TODO: Check CC email is valid
                    if (!String.IsNullOrEmpty(emailCc))
                    {
                        message.CC.Add(emailCc);
                    }
                }
            }
            if (mailBCC != null && mailBCC.Length > 0)
            {
                foreach (string emailBcc in mailBCC)
                {
                    //TODO: Check CC email is valid
                    if (!String.IsNullOrEmpty(emailBcc))
                    {
                        message.Bcc.Add(emailBcc);
                    }
                }
            }
            ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(RemoteServerCertificateValidationCallback);

            var client = new SmtpClient(SMTP_SERVER, SMTP_PORT);


            client.Timeout = 10000;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;

            // Pass SMTP credentials

            // Enable SSL encryption
            client.UseDefaultCredentials = true;
            client.Credentials = new NetworkCredential(SMTP_USERNAME, SMTP_PASSWORD);
            client.EnableSsl = false;
            //client.UseDefaultCredentials = true;
            //client.EnableSsl = false;

            // Try to send the message. Show status in console.

            try
            {
                client.Send(message);
                return "0";
            }
            catch (Exception ex)
            {
                
                string errordt = ex.Message;// MessageBox.Show("Gửi mail " + string.Join("*", mailTo) + ex.Message, "error");
                return errordt;
            }

        }
        /*
        static string[] getcfmail()
        {
            string logfile = System.IO.Directory.GetCurrentDirectory();
            byte[] data = new byte[2049];
            int lenbyte = 0;

            string datastring = "";
            string server = "";
            string name = "";
            string user = "";
            string pass = "";

            string[] add_to_array = new string[4];

            DataCryption crypt = new DataCryption();

            try
            {
                FileStream fs = new FileStream(logfile + "\\data\\mail.cfg", FileMode.OpenOrCreate, FileAccess.Read);
                fs.Read(data, 0, 2048);
                fs.Close();

                for (int i = 0; i < data.Length; i++)
                {
                    if (data[i] == (byte)0)
                        break;
                    lenbyte++;
                }


                byte[] fixdata = new byte[lenbyte];

                Buffer.BlockCopy(data, 0, fixdata, 0, lenbyte);

                string aaaa = Encoding.UTF8.GetString(fixdata);

                string cccc = aaaa.Substring(1, aaaa.Length - 1);

                datastring = Encoding.UTF8.GetString(crypt.decrypt(Encoding.UTF8.GetBytes(cccc)));
                string[] split = datastring.Split('|');

                server = split[0];
                name = split[1];
                user = split[2];
                pass = split[3];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            add_to_array[0] = server;
            add_to_array[1] = name;
            add_to_array[2] = user;
            add_to_array[3] = pass;
            return add_to_array;
        }
        */
        public static bool RemoteServerCertificateValidationCallback(Object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            string AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            AppLocation = AppLocation.Replace("file:\\", "");
            string cer = AppLocation + "\\lib\\3_MA1.cer";
            if (sslPolicyErrors == SslPolicyErrors.None)
                return true;

            // if got an cert auth error
            if (sslPolicyErrors != SslPolicyErrors.RemoteCertificateNameMismatch) return false;
            string sertFileName = cer;

            // check if cert file exists
            if (File.Exists(sertFileName))
            {
                var actualCertificate = X509Certificate.CreateFromCertFile(sertFileName);
                return certificate.Equals(actualCertificate);
            }

            // export and check if cert not exists
            using (var file = File.Create(sertFileName))
            {
                var cert = certificate.Export(X509ContentType.Cert);
                file.Write(cert, 0, cert.Length);
            }

            var createdCertificate = X509Certificate.CreateFromCertFile(sertFileName);
            return certificate.Equals(createdCertificate);
        }
    }
}
