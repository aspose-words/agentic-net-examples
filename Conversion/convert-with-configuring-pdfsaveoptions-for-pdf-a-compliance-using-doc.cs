using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Create a PdfSaveOptions instance to customize PDF output.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Configure the PDF/A compliance level (e.g., PDF/A-1b).
        saveOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the document as a PDF using the configured options.
        doc.Save("Output.pdf", saveOptions);
    }
}
