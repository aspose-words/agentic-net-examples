using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Load the existing DOCX mail‑merge template.
        Document doc = new Document("Template.docx");

        // Customize mail‑merge template properties.
        // Example: set the document type to a form‑letter and the destination to a new document.
        doc.MailMergeSettings.MainDocumentType = MailMergeMainDocumentType.FormLetters;
        doc.MailMergeSettings.Destination = MailMergeDestination.NewDocument;

        // Save the document as a PNG image.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        // Optional: render only the first page (zero‑based index).
        pngOptions.PageSet = new PageSet(0);
        doc.Save("Output.png", pngOptions);
    }
}
