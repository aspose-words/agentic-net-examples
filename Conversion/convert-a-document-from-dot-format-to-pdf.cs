using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOT file
        string inputPath = "Template.dot";

        // Path to the output PDF file
        string outputPath = "Result.pdf";

        // Load the DOT template into a Document object
        Document doc = new Document(inputPath);

        // Create PDF save options (customize as needed)
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        // Example: enable high‑quality rendering for better visual fidelity
        saveOptions.UseHighQualityRendering = true;

        // Save the document as PDF
        doc.Save(outputPath, saveOptions);
    }
}
