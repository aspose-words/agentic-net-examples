using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Saving;

class MailMergeEmailExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert simple MERGEFIELDs that will be filled by the mail merge.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Configure mail‑merge settings so that the result is treated as an e‑mail.
        MailMergeSettings settings = doc.MailMergeSettings;
        settings.Destination = MailMergeDestination.Email;          // Result will be an e‑mail.
        settings.MainDocumentType = MailMergeMainDocumentType.Email; // Optional: specify that the source is an e‑mail type.
        settings.MailSubject = "Personalised greeting";            // Subject line for the e‑mail.
        settings.AddressFieldName = "EmailAddress";                // Column that contains the recipient address.

        // Prepare the data for the mail merge.
        string[] fieldNames = { "FirstName", "LastName", "Message", "EmailAddress" };
        object[] fieldValues = { "John", "Doe", "Hello! This message was created with Aspose.Words mail merge.", "john.doe@example.com" };

        // Execute the mail merge. The result will be stored in the document.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document to a memory stream in MHTML format.
        // MHTML contains the full e‑mail body (HTML + resources) which can be used as the e‑mail content.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            doc.Save(mhtmlStream, SaveFormat.Mhtml);
            mhtmlStream.Position = 0; // Reset stream position for reading.

            // Read the MHTML content as a string.
            string mhtmlContent;
            using (StreamReader reader = new StreamReader(mhtmlStream))
                mhtmlContent = reader.ReadToEnd();

            // Create a MailMessage and populate its fields.
            MailMessage message = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = settings.MailSubject,
                // The MHTML content includes the required MIME headers; for simplicity we use it as the body.
                Body = mhtmlContent,
                IsBodyHtml = true
            };
            // Add the recipient address taken from the merged data.
            message.To.Add(new MailAddress("john.doe@example.com"));

            // Optionally attach the merged document as a .docx file.
            using (MemoryStream docxStream = new MemoryStream())
            {
                doc.Save(docxStream, SaveFormat.Docx);
                docxStream.Position = 0;
                Attachment attachment = new Attachment(docxStream, "MergedDocument.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
                message.Attachments.Add(attachment);

                // Send the e‑mail using SmtpClient (configure host/credentials as needed).
                using (SmtpClient client = new SmtpClient("smtp.example.com", 587))
                {
                    client.Credentials = new NetworkCredential("smtp_user", "smtp_password");
                    client.EnableSsl = true;

                    // Uncomment the line below to actually send the e‑mail.
                    // client.Send(message);
                }
            }
        }
    }
}
