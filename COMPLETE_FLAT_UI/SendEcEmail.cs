using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net.Mail;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Pkcs;
using System.Net;

namespace MAS_EMAIL
{
     class SendEcEmail
     {
        public string AppLocation = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
        public string SendEmail(string[] mailTo, string[] mailCC, string[] mailBCC, string subject, string mailBody, string[] attachments, string acnt_no, bool ismailBodyHtml)
        {
            //set my certificate info up
            string CertificatePath = AppLocation+ "\\masvn.pfx";
            string CertificatePassword = "123456";
            string MailServer = Common.SMTP_SERVER;
            int SMTP_PORT = Int32.Parse(Common.SMTP_PORT);
            string Emailsendnm = Common.MAIL_FROM_NAME;
            string SMTP_USERNAME = Common.SMTP_USERNAME;
            string SMTP_PASSWORD = Common.SMTP_PASSWORD;


            String SUBJECT = subject;
            // The body of the email
            String BODY = mailBody;

            string EmailSender = SMTP_USERNAME;
            bool IsHtmlEmail = ismailBodyHtml;

            //Load the certificate
            X509Certificate2 EncryptCert =
               new X509Certificate2(CertificatePath, CertificatePassword);

            //Build the body into a string
            StringBuilder Message = new StringBuilder();
            Message.AppendLine("Content-Type: text/" +
                ((IsHtmlEmail) ? "html" : "plain") +
                "; charset=\"UTF-8\"");

            Message.AppendLine("Content-Transfer-Encoding: 7bit");
            Message.AppendLine(BODY);

            //Convert the body to bytes
            byte[] BodyBytes = Encoding.ASCII.GetBytes(Message.ToString());

            //Build the e-mail body bytes into a secure envelope
            EnvelopedCms Envelope = new EnvelopedCms(new ContentInfo(BodyBytes));
            CmsRecipient Recipient = new CmsRecipient(
                SubjectIdentifierType.IssuerAndSerialNumber, EncryptCert);
            Envelope.Encrypt(Recipient);
            byte[] EncryptedBytes = Envelope.Encode();

            //Creat the mail message
            MailMessage Msg = new MailMessage();
            //Msg.To.Add(new MailAddress(EmailRecipient));
            Msg.From = new MailAddress(EmailSender, Emailsendnm);
            Msg.Subject = SUBJECT;
           // string htmlText = BODY;

            Msg.IsBodyHtml = true;

            if (attachments != null && attachments.Length > 0)
            {
                foreach (string attachmentsen in attachments)
                {
                    //TODO: Check CC email is valid
                    if (!String.IsNullOrEmpty(attachmentsen))
                    {
                        Msg.Attachments.Add(new System.Net.Mail.Attachment(attachmentsen));
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
                        Msg.To.Add(new MailAddress(emailTo));
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
                        Msg.CC.Add(emailCc);
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
                        Msg.Bcc.Add(emailBcc);
                    }
                }
            }
            //AlternateView plainView = AlternateView.CreateAlternateViewFromString("Some plaintext", Encoding.UTF8, "text/plain");
            // We have something to show in real old mail clients. 
            // AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlText, Encoding.UTF8, "text/html");

            //Attach the encrypted body to the email as and ALTERNATE VIEW
            MemoryStream ms = new MemoryStream(EncryptedBytes);
            AlternateView av = new AlternateView(ms,"application/pkcs7-mime; smime-type=signed-data;name=smime.p7m");
            Msg.AlternateViews.Add(av);

            SmtpClient smtp = new SmtpClient(MailServer, SMTP_PORT);
            smtp.Credentials = new NetworkCredential(SMTP_USERNAME, SMTP_PASSWORD);
            smtp.UseDefaultCredentials = true;
            smtp.EnableSsl = true;
            //send the email    

            try
            {
                smtp.Send(Msg);
            }
            catch (Exception ex)
            {

                string errordt = ex.Message;// MessageBox.Show("Gửi mail " + string.Join("*", mailTo) + ex.Message, "error");
                return errordt;
            }
            return "0";
        }
     }
}
