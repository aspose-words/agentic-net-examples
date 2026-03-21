using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfSubsetExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a simple DOCX document in memory.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello, Aspose.Words PDF subsetting example!");

            // Determine output path (current directory).
            string outputDirectory = Directory.GetCurrentDirectory();
            string pdfPath = Path.Combine(outputDirectory, "SampleDocument_Subset.pdf");

            // Configure PDF save options to embed only used glyphs.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Subsetting is enabled when EmbedFullFonts is false.
                EmbedFullFonts = false,
                // Embed all fonts that are used in the document.
                FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
            };

            // Save the document as PDF using the configured options.
            doc.Save(pdfPath, pdfOptions);

            Console.WriteLine($"PDF saved to: {pdfPath}");
        }
    }
}
