using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source PDF/A or PDF/UA document.
        string inputPath = "input.pdf";

        // Path where the resulting PDF will be saved.
        string outputPath = "output.pdf";

        // Load the existing PDF document.
        Document doc = new Document(inputPath);

        // Example modification: insert a new paragraph at the end of the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This text was added after loading the PDF.");

        // Create a PdfSaveOptions object to specify PDF save settings.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the desired PDF compliance level (e.g., PDF/A-2u). Change as needed.
        saveOptions.Compliance = PdfCompliance.PdfA2u;

        // Save the document as PDF using the specified compliance options.
        doc.Save(outputPath, saveOptions);
    }
}
