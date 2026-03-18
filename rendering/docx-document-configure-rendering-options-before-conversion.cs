using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentConversion
{
    static void Main()
    {
        // Path to the source DOCX file.
        const string inputPath = @"C:\Docs\InputDocument.docx";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // --------------------------------------------------------------------
        // Configure rendering options that affect how the document is laid out
        // before it is converted to another format.
        // --------------------------------------------------------------------
        // Example: Show hidden text in the rendered output.
        doc.LayoutOptions.ShowHiddenText = true;

        // Example: Do not display paragraph marks in the rendered output.
        doc.LayoutOptions.ShowParagraphMarks = false;

        // Example: Set revision display options (optional).
        doc.LayoutOptions.RevisionOptions.ShowRevisionBars = true;
        // The RevisionColor enum is not available in the current Aspose.Words version.
        // If you need to colour inserted text, use the InsertedTextColor property with a supported enum value or omit it.
        // doc.LayoutOptions.RevisionOptions.InsertedTextColor = RevisionColor.BrightGreen; // Removed for compatibility

        // --------------------------------------------------------------------
        // Configure save options for the target format (PDF in this case).
        // --------------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render DrawingML shapes using their fallback shapes.
            DmlRenderingMode = DmlRenderingMode.Fallback,

            // Enable memory optimization for large documents.
            MemoryOptimization = true,

            // Optional: Set the PDF compliance level.
            Compliance = PdfCompliance.PdfA1b
        };

        // Path to the output PDF file.
        const string outputPath = @"C:\Docs\ConvertedDocument.pdf";

        // Save the document to PDF using the configured options.
        doc.Save(outputPath, pdfOptions);
    }
}
