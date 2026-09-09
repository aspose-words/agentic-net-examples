using System;
using System.IO;
using System.Net.Mail;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words! This document will be converted to MHTML and embedded in an email.");
        const string docxPath = "sample.docx";
        doc.Save(docxPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(docxPath) || new FileInfo(docxPath).Length == 0)
            throw new InvalidOperationException("Failed to create the sample DOCX file.");

        // Step 2: Load the DOCX and convert it to MHTML.
        Document loadedDoc = new Document(docxPath);
        const string mhtmlPath = "sample.mht";
        loadedDoc.Save(mhtmlPath, SaveFormat.Mhtml);

        // Verify that the MHTML file was created.
        if (!File.Exists(mhtmlPath) || new FileInfo(mhtmlPath).Length == 0)
            throw new InvalidOperationException("Failed to convert the document to MHTML.");

        // Read the MHTML content.
        string mhtmlContent = File.ReadAllText(mhtmlPath);

        // Step 3: Create a simple email and embed the MHTML content as the HTML body.
        // Using System.Net.Mail instead of Aspose.Email (which is not part of the required packages).
        MailMessage email = new MailMessage
        {
            From = new MailAddress("sender@example.com"),
            Subject = "Document embedded as MHTML",
            IsBodyHtml = true,
            Body = mhtmlContent
        };
        email.To.Add("recipient@example.com");

        // Save the email content to an .eml file for verification.
        const string emlPath = "email.eml";
        File.WriteAllText(emlPath, email.ToString());

        // Verify that the email file was created.
        if (!File.Exists(emlPath) || new FileInfo(emlPath).Length == 0)
            throw new InvalidOperationException("Failed to save the email message.");

        // Output simple confirmation.
        Console.WriteLine("DOCX to MHTML conversion and email embedding completed successfully.");
    }
}
