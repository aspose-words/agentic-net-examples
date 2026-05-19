using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a simple document that uses a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Common TrueType font.
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");

        // Enable embedding of TrueType fonts.
        doc.FontInfos.EmbedTrueTypeFonts = true;

        // -----------------------------------------------------------------
        // Render TIFF with full (non‑subsetted) embedded fonts.
        // -----------------------------------------------------------------
        doc.FontInfos.SaveSubsetFonts = false; // Do not subset fonts.
        ImageSaveOptions fullOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300
        };
        string fullTiffPath = Path.Combine(outputDir, "Document_FullFonts.tiff");
        doc.Save(fullTiffPath, fullOptions);

        // -----------------------------------------------------------------
        // Render TIFF with subsetted embedded fonts.
        // -----------------------------------------------------------------
        // NOTE: Font subsetting is not supported for image formats (TIFF, PNG, etc.).
        // The SaveSubsetFonts property only affects DOC/DOCX/RTF saving.
        // Therefore the file sizes may be identical. We still perform the operation
        // to demonstrate the API usage without throwing an exception.
        doc.FontInfos.SaveSubsetFonts = true; // Request subsetting (no effect for TIFF).
        ImageSaveOptions subsetOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300
        };
        string subsetTiffPath = Path.Combine(outputDir, "Document_SubsetFonts.tiff");
        doc.Save(subsetTiffPath, subsetOptions);

        // Verify that both files were created.
        if (!File.Exists(fullTiffPath) || !File.Exists(subsetTiffPath))
            throw new InvalidOperationException("One or more TIFF files were not created.");

        // Compare file sizes.
        long fullSize = new FileInfo(fullTiffPath).Length;
        long subsetSize = new FileInfo(subsetTiffPath).Length;

        Console.WriteLine($"Full‑font TIFF size   : {fullSize} bytes");
        Console.WriteLine($"Subset‑font TIFF size : {subsetSize} bytes");

        // If subsetting had an effect, the subset file would be smaller.
        // For TIFF this is not the case, so we only warn instead of throwing.
        if (subsetSize >= fullSize)
        {
            Console.WriteLine("Warning: Subsetting did not reduce the TIFF file size. " +
                              "Font subsetting is not supported for image formats.");
        }
    }
}
