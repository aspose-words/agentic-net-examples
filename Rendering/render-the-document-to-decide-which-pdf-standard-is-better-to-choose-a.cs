using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfStandardSelector
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = "input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = "output.pdf";

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Create PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Decide which PDF compliance level to use.
        // Example logic: if the document contains any images, use PDF/A-1b for archival purposes;
        // otherwise, use the default PDF 1.7 compliance.
        bool containsImages = doc.GetChildNodes(NodeType.Shape, true).Count > 0;

        if (containsImages)
        {
            // Preserve visual appearance for documents with images.
            pdfOptions.Compliance = PdfCompliance.PdfA1b;
        }
        else
        {
            // Use the standard PDF 1.7 compliance.
            pdfOptions.Compliance = PdfCompliance.Pdf17;
        }

        // Save the document as PDF using the selected compliance level.
        doc.Save(outputPath, pdfOptions);
    }
}
