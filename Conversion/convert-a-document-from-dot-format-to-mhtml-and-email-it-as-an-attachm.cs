using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

// Entry point for the console application.
class Program
{
    static void Main()
    {
        var emailer = new DotToMhtmlEmailer();
        emailer.ConvertAndSend(
            dotFilePath: @"C:\Docs\Template.dot",
            mhtmlFilePath: @"C:\Docs\Output.mhtml",
            smtpHost: "smtp.example.com",
            smtpPort: 587,
            fromAddress: "sender@example.com",
            toAddress: "recipient@example.com",
            subject: "Converted Document",
            body: "Please find the converted MHTML document attached.",
            smtpUser: "smtp_user",
            smtpPassword: "smtp_password");
    }
}

public class DotToMhtmlEmailer
{
    /// <summary>
    /// Converts a .dot template to .mhtml and sends it as an email attachment.
    /// </summary>
    public void ConvertAndSend(string dotFilePath, string mhtmlFilePath,
                               string smtpHost, int smtpPort,
                               string fromAddress, string toAddress,
                               string subject, string body,
                               string? smtpUser = null, string? smtpPassword = null)
    {
        // Load the DOT document.
        Document doc = new Document(dotFilePath);

        // Save the document as MHTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        doc.Save(mhtmlFilePath, saveOptions);

        // Prepare and send the email.
        using (MailMessage message = new MailMessage())
        {
            message.From = new MailAddress(fromAddress);
            message.To.Add(new MailAddress(toAddress));
            message.Subject = subject;
            message.Body = body;
            message.IsBodyHtml = false;

            // Attach the generated MHTML file.
            message.Attachments.Add(new Attachment(mhtmlFilePath));

            using (SmtpClient smtpClient = new SmtpClient(smtpHost, smtpPort))
            {
                smtpClient.EnableSsl = true;

                if (!string.IsNullOrEmpty(smtpUser) && !string.IsNullOrEmpty(smtpPassword))
                {
                    smtpClient.Credentials = new NetworkCredential(smtpUser, smtpPassword);
                }
                else
                {
                    smtpClient.UseDefaultCredentials = true;
                }

                smtpClient.Send(message);
            }
        }

        // Optional: delete the temporary MHTML file if you no longer need it.
        // File.Delete(mhtmlFilePath);
    }
}
