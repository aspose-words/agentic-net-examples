using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Settings;

class MailMergeTemplateExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Add a static header that will appear on every merged document -----
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Company Confidential Report");
        builder.Writeln(); // Add an empty line after the header.

        // Reset formatting for the body of the template.
        builder.Font.Size = 12;
        builder.Font.Bold = false;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // ----- Insert mail merge fields -----
        // Example fields: FirstName, LastName, Address
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.Writeln();
        builder.Write("Address: ");
        builder.InsertField("MERGEFIELD Address", "<Address>");
        builder.Writeln();

        // ----- Configure MailMergeSettings (optional) -----
        // Set the main document type to FormLetters so Word knows this is a mail‑merge template.
        MailMergeSettings settings = doc.MailMergeSettings;
        settings.MainDocumentType = MailMergeMainDocumentType.FormLetters;

        // Save the template to a file.
        doc.Save("MailMergeTemplate.docx");
    }
}
