using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Document.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\Output.pdf";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Configure PDF save options to embed full TrueType fonts.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // When true, the complete font file (all glyphs) is embedded in the PDF.
            EmbedFullFonts = true
        };

        // Save the document as PDF using the configured options.
        doc.Save(outputPath, pdfOptions);
    }
}
