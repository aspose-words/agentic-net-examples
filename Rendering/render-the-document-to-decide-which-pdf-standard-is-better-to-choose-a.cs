using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing; // Added for Shape class

class Program
{
    static void Main()
    {
        // Load the source document.
        // Replace with the actual path to your DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";
        Document doc = new Document(inputPath);

        // Decide which PDF compliance level to use.
        // Example logic: if the document contains any images, use PDF/A-1b for archival purposes;
        // otherwise, use the default PDF 1.7 compliance.
        bool containsImages = false;
        foreach (Node node in doc.GetChildNodes(NodeType.Shape, true))
        {
            Shape shape = (Shape)node;
            if (shape.HasImage)
            {
                containsImages = true;
                break;
            }
        }

        // Create PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the compliance level based on the decision.
        pdfOptions.Compliance = containsImages
            ? PdfCompliance.PdfA1b   // Preserve visual appearance for documents with images.
            : PdfCompliance.Pdf17;   // Standard PDF 1.7 for other cases.

        // Save the document as PDF using the configured options.
        // Replace with the desired output path.
        string outputPath = @"C:\Docs\ResultDocument.pdf";
        doc.Save(outputPath, pdfOptions);
    }
}
