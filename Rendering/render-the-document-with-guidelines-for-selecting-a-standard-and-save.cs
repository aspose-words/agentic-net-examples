using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document that contains the guidelines.
        string inputPath = "Guidelines.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = "Guidelines.pdf";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Create PDF save options and set the desired PDF compliance level.
        // Here we choose PDF/A-1b which preserves visual appearance.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the document as a PDF using the specified options.
        doc.Save(outputPath, saveOptions);
    }
}
