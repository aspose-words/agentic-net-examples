using System;
using Aspose.Words;
using Aspose.Words.Settings;

namespace MailMergeTemplateDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add merge fields to the template.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Example merge fields for a typical letter.
            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD LastName", "<LastName>");
            builder.Writeln(":");
            builder.InsertField("MERGEFIELD Message", "<Message>");

            // Optionally configure mail merge settings (e.g., specify that this is a form‑letter document).
            MailMergeSettings settings = doc.MailMergeSettings;
            settings.MainDocumentType = MailMergeMainDocumentType.FormLetters;

            // Save the template as a DOCX file.
            doc.Save("MailMergeTemplate.docx");
        }
    }
}
