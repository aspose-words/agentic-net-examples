using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing document (replace with your source file path)
        Document doc = new Document("input.docx");

        // Configure save options for DOCX format
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            // Explicitly set the target format (optional, default is Docx)
            SaveFormat = SaveFormat.Docx,

            // Example of additional option: enable pretty formatting for readability
            PrettyFormat = true,

            // Example: embed the generator name (default is true)
            ExportGeneratorName = true
        };

        // Save the document using the configured options
        doc.Save("output.docx", saveOptions);
    }
}
