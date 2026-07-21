using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Helper method to create a sample document with some text.
        Document CreateSampleDocument()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Name = "Arial";
            builder.Font.Size = 24;
            builder.Writeln("The quick brown fox jumps over the lazy dog. 0123456789");
            builder.Writeln("ABCDEFGHIJKLMNOPQRSTUVWXYZ");
            builder.Writeln("abcdefghijklmnopqrstuvwxyz");
            return doc;
        }

        // -----------------------------------------------------------------
        // Full‑font PDF (no subsetting)
        // -----------------------------------------------------------------
        Document docFull = CreateSampleDocument();

        // Enable embedding of TrueType fonts and disable subsetting.
        docFull.FontInfos.EmbedTrueTypeFonts = true;
        docFull.FontInfos.SaveSubsetFonts = false; // embed the whole font

        string pdfFullPath = Path.Combine(outputDir, "FullFont.pdf");
        PdfSaveOptions fullOptions = new PdfSaveOptions
        {
            // When EmbedFullFonts is true the whole font file is embedded.
            EmbedFullFonts = true
        };
        docFull.Save(pdfFullPath, fullOptions);

        // -----------------------------------------------------------------
        // Subset‑font PDF (subsetting enabled)
        // -----------------------------------------------------------------
        Document docSubset = CreateSampleDocument();

        // Enable embedding of TrueType fonts and enable subsetting.
        docSubset.FontInfos.EmbedTrueTypeFonts = true;
        docSubset.FontInfos.SaveSubsetFonts = true; // embed only used glyphs

        string pdfSubsetPath = Path.Combine(outputDir, "SubsetFont.pdf");
        PdfSaveOptions subsetOptions = new PdfSaveOptions
        {
            // When EmbedFullFonts is false the fonts are subsetted before embedding.
            EmbedFullFonts = false
        };
        docSubset.Save(pdfSubsetPath, subsetOptions);

        // Verify that both PDF files were created.
        if (!File.Exists(pdfFullPath) || !File.Exists(pdfSubsetPath))
            throw new FileNotFoundException("One of the PDF files was not created.");

        // Compare file sizes.
        long fullSize = new FileInfo(pdfFullPath).Length;
        long subsetSize = new FileInfo(pdfSubsetPath).Length;

        Console.WriteLine($"Full‑font PDF size   : {fullSize} bytes");
        Console.WriteLine($"Subset‑font PDF size : {subsetSize} bytes");

        if (subsetSize >= fullSize)
            throw new InvalidOperationException("Subsetting did not reduce the PDF file size as expected.");
    }
}
