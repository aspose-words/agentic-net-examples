using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Create a SaveOptions object suitable for PDF format.
        // This uses the provided SaveOptions.CreateSaveOptions(SaveFormat) rule.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Example configuration: enable memory optimization for large documents.
        saveOptions.MemoryOptimization = true;

        // Save the document as PDF using the configured SaveOptions.
        // This uses the Document.Save(string, SaveOptions) rule.
        doc.Save("Output.pdf", saveOptions);
    }
}
