using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create a sample document with several pages of text using a TrueType font.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Font.Name = "Arial"; // Arial is a TrueType font available on most systems.
        builder.Font.Size = 24;

        // Add content for 5 pages.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}. The quick brown fox jumps over the lazy dog.");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // -----------------------------------------------------------------
        // Render to PDF with full font embedding (no subsetting).
        // -----------------------------------------------------------------
        Document fullEmbeddingDoc = (Document)sourceDoc.Clone(true);
        fullEmbeddingDoc.FontInfos.EmbedTrueTypeFonts = true;   // Enable embedding.
        fullEmbeddingDoc.FontInfos.SaveSubsetFonts = false;    // Do NOT subset.

        string fullPdfPath = Path.Combine(outputDir, "FullEmbedding.pdf");
        PdfSaveOptions fullPdfOptions = new PdfSaveOptions
        {
            // Embed the complete font (no subsetting).
            EmbedFullFonts = true
        };
        fullEmbeddingDoc.Save(fullPdfPath, fullPdfOptions);

        // -----------------------------------------------------------------
        // Render to PDF with font subsetting enabled.
        // -----------------------------------------------------------------
        Document subsetEmbeddingDoc = (Document)sourceDoc.Clone(true);
        subsetEmbeddingDoc.FontInfos.EmbedTrueTypeFonts = true; // Enable embedding.
        subsetEmbeddingDoc.FontInfos.SaveSubsetFonts = true;    // Enable subsetting.

        string subsetPdfPath = Path.Combine(outputDir, "SubsetEmbedding.pdf");
        PdfSaveOptions subsetPdfOptions = new PdfSaveOptions
        {
            // Default behavior (EmbedFullFonts = false) will subset the fonts.
            EmbedFullFonts = false
        };
        subsetEmbeddingDoc.Save(subsetPdfPath, subsetPdfOptions);

        // -----------------------------------------------------------------
        // Validate that both files were created and compare their sizes.
        // -----------------------------------------------------------------
        if (!File.Exists(fullPdfPath) || !File.Exists(subsetPdfPath))
            throw new FileNotFoundException("One of the PDF files was not created.");

        long fullSize = new FileInfo(fullPdfPath).Length;
        long subsetSize = new FileInfo(subsetPdfPath).Length;

        Console.WriteLine($"Full embedding PDF size   : {fullSize} bytes");
        Console.WriteLine($"Subset embedding PDF size : {subsetSize} bytes");

        // Ensure that subsetting actually reduced the file size.
        if (subsetSize >= fullSize)
            throw new InvalidOperationException("Subsetting did not reduce the PDF file size as expected.");
    }
}
