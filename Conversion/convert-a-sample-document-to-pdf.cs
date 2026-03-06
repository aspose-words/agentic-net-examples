using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = "SampleDocument.docx";

        // Path where the PDF will be saved.
        string outputPath = "SampleDocument.pdf";

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Create PDF-specific save options.
        SaveOptions pdfOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Save the loaded document as a PDF file.
        doc.Save(outputPath, pdfOptions);
    }
}
