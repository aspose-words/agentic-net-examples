using System;
using Aspose.Words;
using Aspose.Words.Saving;

class FontEmbeddingPdfExample
{
    static void Main()
    {
        // Load an existing document or create a new one.
        // Here we create a new document for demonstration.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text with custom font settings.");

        // Prepare PDF save options.
        PdfSaveOptions options = new PdfSaveOptions();

        // Specify the desired font embedding mode.
        // PdfFontEmbeddingMode can be EmbedAll, EmbedNonstandard, or EmbedNone.
        // Replace 'pdfFontEmbeddingMode' with the required enum value.
        PdfFontEmbeddingMode pdfFontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll;
        options.FontEmbeddingMode = pdfFontEmbeddingMode;

        // Optionally control whether fonts are fully embedded (no subsetting).
        // Set to true to embed the complete font file.
        bool embedFullFonts = false;
        options.EmbedFullFonts = embedFullFonts;

        // Save the document as PDF with the configured font embedding options.
        string outputPath = "RenderedDocument.pdf";
        doc.Save(outputPath, options);
    }
}
