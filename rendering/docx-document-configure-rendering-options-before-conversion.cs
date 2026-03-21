using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentConversion
{
    static void Main()
    {
        // Create a new document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

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

        // Path to the output PDF file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConvertedDocument.pdf");

        // Save the document to PDF using the configured options.
        doc.Save(outputPath, pdfOptions);

        Console.WriteLine($"Document successfully saved to: {outputPath}");
    }
}
