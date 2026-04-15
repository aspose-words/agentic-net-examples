using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;   // Needed for FontInfoCollection

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the two TIFF files.
        string tiffFullPath = Path.Combine(outputDir, "Document_FullFonts.tiff");
        string tiffSubsetPath = Path.Combine(outputDir, "Document_SubsetFonts.tiff");

        // Create a sample document with several pages of text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";               // Use a TrueType font.
        builder.Font.Size = 12;

        for (int page = 1; page <= 5; page++)
        {
            builder.Writeln($"Page {page}");
            // Add enough text to make the file size noticeable.
            for (int i = 0; i < 100; i++)
            {
                builder.Writeln("The quick brown fox jumps over the lazy dog. 1234567890");
            }

            if (page < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure font embedding: embed TrueType fonts.
        FontInfoCollection fontInfos = doc.FontInfos;
        fontInfos.EmbedTrueTypeFonts = true;

        // Prepare image save options for TIFF.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use a reasonable resolution; default compression (LZW) is fine.
            Resolution = 150
        };

        // ---------- Save with full font embedding (no subsetting) ----------
        fontInfos.SaveSubsetFonts = false; // Save the whole font.
        doc.Save(tiffFullPath, tiffOptions);

        // ---------- Save with subset font embedding (subsetting enabled) ----------
        // Clone the original document to keep the first file unchanged.
        Document docSubset = (Document)doc.Clone(true);
        FontInfoCollection subsetFontInfos = docSubset.FontInfos;
        subsetFontInfos.EmbedTrueTypeFonts = true;
        subsetFontInfos.SaveSubsetFonts = true; // Enable subsetting.
        docSubset.Save(tiffSubsetPath, tiffOptions);

        // Validate that both files were created.
        if (!File.Exists(tiffFullPath))
            throw new FileNotFoundException("Full‑font TIFF was not created.", tiffFullPath);
        if (!File.Exists(tiffSubsetPath))
            throw new FileNotFoundException("Subset‑font TIFF was not created.", tiffSubsetPath);

        // Compare file sizes.
        long fullSize = new FileInfo(tiffFullPath).Length;
        long subsetSize = new FileInfo(tiffSubsetPath).Length;

        Console.WriteLine($"Full‑font TIFF size   : {fullSize} bytes");
        Console.WriteLine($"Subset‑font TIFF size : {subsetSize} bytes");
        Console.WriteLine($"Size reduction       : {fullSize - subsetSize} bytes");
    }
}
