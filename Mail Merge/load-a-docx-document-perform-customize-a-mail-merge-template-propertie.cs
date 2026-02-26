using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

class MailMergeToPng
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\Template.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Customize mail‑merge template properties.
        // Set the document type to a form letter. This tells Aspose.Words that the
        // document will be used as a mail‑merge template.
        doc.MailMergeSettings.MainDocumentType = MailMergeMainDocumentType.FormLetters;

        // (Optional) Perform a simple mail merge to populate fields.
        // This step is not required for just customizing the template,
        // but demonstrates that the settings are applied.
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Prepare image save options for PNG format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render the first page only; change as needed.
            PageSet = new PageSet(0),

            // Set a reasonable resolution (dpi).
            Resolution = 300
        };

        // Path to the output PNG file.
        string outputPath = @"C:\Docs\Result.png";

        // Save the document as a PNG image.
        doc.Save(outputPath, saveOptions);
    }
}
