using System;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Load the DOCX file that will be used as a mail‑merge template.
        Document doc = new Document("Template.docx");

        // Access the mail‑merge settings of the loaded document.
        // This demonstrates how to work with a mail‑merge template.
        MailMergeSettings settings = doc.MailMergeSettings;

        // Example: explicitly mark the document as not being a mail‑merge document.
        // (Adjust as needed for your scenario.)
        settings.MainDocumentType = MailMergeMainDocumentType.NotAMergeDocument;

        // Render the first page of the document to a PNG image and save it.
        doc.Save("TemplatePage.png", SaveFormat.Png);
    }
}
