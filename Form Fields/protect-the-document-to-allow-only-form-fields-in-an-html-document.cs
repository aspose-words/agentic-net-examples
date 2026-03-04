using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ProtectDocumentForFormFields
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = "input.docx";

        // Path where the protected HTML will be saved.
        string outputPath = "output.html";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Protect the document so that only form fields can be edited.
        doc.Protect(ProtectionType.AllowOnlyFormFields);

        // Configure HTML save options to export form fields as interactive HTML input elements.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportFormFields = true
        };

        // Save the protected document as HTML.
        doc.Save(outputPath, htmlOptions);
    }
}
