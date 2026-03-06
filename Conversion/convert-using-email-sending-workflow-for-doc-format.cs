using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

public class DocumentEmailSender
{
    public void SendDocumentAsDoc(string inputFilePath,
                                  string smtpHost,
                                  int smtpPort,
                                  string smtpUser,
                                  string smtpPassword,
                                  string fromAddress,
                                  string toAddress,
                                  string subject,
                                  string body)
    {
        // Load the source document.
        Document doc = new Document(inputFilePath);

        // Create a temporary file for the DOC output.
        string tempDocPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".doc");
        doc.Save(tempDocPath, SaveFormat.Doc);

        // Build the e‑mail message.
        using (MailMessage message = new MailMessage())
        {
            message.From = new MailAddress(fromAddress);
            message.To.Add(new MailAddress(toAddress));
            message.Subject = subject;
            message.Body = body;
            message.IsBodyHtml = false;

            // Attach the converted DOC file.
            using (Attachment attachment = new Attachment(tempDocPath))
            {
                message.Attachments.Add(attachment);

                // Configure and send via SMTP.
                using (SmtpClient client = new SmtpClient(smtpHost, smtpPort))
                {
                    client.EnableSsl = true;
                    if (!string.IsNullOrEmpty(smtpUser))
                    {
                        client.Credentials = new NetworkCredential(smtpUser, smtpPassword);
                    }
                    else
                    {
                        client.UseDefaultCredentials = true;
                    }
                    client.Send(message);
                }
            }
        }

        // Delete the temporary file.
        if (File.Exists(tempDocPath))
        {
            File.Delete(tempDocPath);
        }
    }
}

public static class Program
{
    public static void Main(string[] args)
    {
        // Expected arguments:
        // 0 – inputFilePath
        // 1 – smtpHost
        // 2 – smtpPort
        // 3 – smtpUser (optional, use empty string if not needed)
        // 4 – smtpPassword (optional)
        // 5 – fromAddress
        // 6 – toAddress
        // 7 – subject
        // 8 – body
        if (args.Length < 9)
        {
            Console.WriteLine("Usage: <inputFilePath> <smtpHost> <smtpPort> <smtpUser> <smtpPassword> <fromAddress> <toAddress> <subject> <body>");
            return;
        }

        var sender = new DocumentEmailSender();
        sender.SendDocumentAsDoc(
            inputFilePath: args[0],
            smtpHost: args[1],
            smtpPort: int.Parse(args[2]),
            smtpUser: args[3],
            smtpPassword: args[4],
            fromAddress: args[5],
            toAddress: args[6],
            subject: args[7],
            body: args[8]);
    }
}
