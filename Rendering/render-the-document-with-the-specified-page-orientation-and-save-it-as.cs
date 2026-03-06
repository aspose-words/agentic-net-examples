using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document.
        Document doc = new Document("Input.docx");

        // Set the desired page orientation for every section in the document.
        // Change Orientation.Landscape to Orientation.Portrait for portrait mode.
        foreach (Section section in doc.Sections)
        {
            section.PageSetup.Orientation = Orientation.Landscape;
        }

        // Rebuild the internal page layout after modifying the orientation.
        doc.UpdatePageLayout();

        // Create PDF save options. Additional options (e.g., PageLayout) can be set here if needed.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Example: display two pages side‑by‑side when the PDF is opened.
        // pdfOptions.PageLayout = PdfPageLayout.TwoPageLeft;

        // Save the document as a PDF file using the specified options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
