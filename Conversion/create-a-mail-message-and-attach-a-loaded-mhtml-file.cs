using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the MHTML file into an Aspose.Words Document.
        // The load operation follows the provided load rule.
        Document mhtmlDoc = new Document("input.mhtml");

        // Prepare a memory stream to hold the MHTML content when saved.
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            // Save the document back to MHTML format into the stream.
            // The save operation follows the provided save rule.
            mhtmlDoc.Save(mhtmlStream, SaveFormat.Mhtml);
            mhtmlStream.Position = 0; // Reset stream position for reading.

            // Create the mail message.
            MailMessage mail = new MailMessage
            {
                From = new MailAddress("sender@example.com"),
                Subject = "Here is the MHTML document",
                Body = "Please find the attached MHTML file."
            };
            mail.To.Add("recipient@example.com");

            // Attach the MHTML content as a file attachment.
            // The attachment uses the stream containing the saved MHTML.
            Attachment attachment = new Attachment(mhtmlStream, "document.mhtml", "message/rfc822");
            mail.Attachments.Add(attachment);

            // At this point the mail message is ready to be sent.
            // For demonstration purposes we will just output a confirmation.
            Console.WriteLine("Mail message created with MHTML attachment.");
        }
    }
}
