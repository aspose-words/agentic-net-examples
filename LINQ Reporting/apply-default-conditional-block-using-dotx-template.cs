using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the template (DOTX) and the output document.
        string templatePath = @"C:\Docs\Business brochure.dotx";
        string outputPath   = @"C:\Docs\Document.DefaultTemplate.docx";

        // Create a new blank document.
        Document doc = new Document();

        // Enable automatic style updating so that the template's styles are applied
        // when the document is opened in Microsoft Word.
        doc.AutomaticallyUpdateStyles = true;

        // Configure SaveOptions to use the specified DOTX as the default template.
        // This will be applied only if the document does not already have an attached template.
        SaveOptions options = SaveOptions.CreateSaveOptions(outputPath);
        options.DefaultTemplate = templatePath;

        // Save the document using the configured options.
        doc.Save(outputPath, options);
    }
}
