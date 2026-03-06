using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the folder that contains the source document.
        string dataDir = @"C:\Data\";

        // Load the existing Word document that contains the guidelines.
        Document doc = new Document(Path.Combine(dataDir, "Guidelines.docx"));

        // Create PDF save options and specify the desired PDF standard compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: comply with PDF/A-1b (ISO 19005-1) to preserve visual appearance.
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as a PDF using the configured options.
        doc.Save(Path.Combine(dataDir, "Guidelines.pdf"), pdfOptions);
    }
}
