using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Create a SaveOptions object suitable for PDF format.
        SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Configure an option – for example, enable memory optimization.
        saveOptions.MemoryOptimization = true;

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", saveOptions);
    }
}
