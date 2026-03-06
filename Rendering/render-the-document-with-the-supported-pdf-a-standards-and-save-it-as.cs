using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Create PDF save options and set the desired PDF/A compliance level.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA2u; // Example: PDF/A-2u compliance.

        // Save the document as a PDF using the specified compliance.
        doc.Save("Output.pdf", saveOptions);
    }
}
