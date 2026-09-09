using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string tiffPath = Path.Combine(outputDir, "Ligatures.tiff");

        // Build a simple document containing characters that form ligatures.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";      // Arial includes common ligatures.
        builder.Font.Size = 48;
        builder.Writeln("Office");        // Contains "ff".
        builder.Writeln("Affix");         // Contains "fi".
        builder.Writeln("Fluff");         // Contains "fl".

        // If custom fonts are required, configure FontSettings here.
        // FontSettings fontSettings = new FontSettings();
        // fontSettings.SetFontsFolder(@"C:\MyFonts", true);
        // doc.FontSettings = fontSettings;

        // Set up TIFF rendering options.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.Resolution = 300;                 // 300 DPI for good quality.
        options.UseAntiAliasing = true;           // Enable anti‑aliasing.
        options.UseHighQualityRendering = true;   // Use high‑quality rendering.

        // Render all pages. Use an explicit int array to avoid ambiguity with the PageSet constructors.
        options.PageSet = new PageSet(new int[] { 0 });

        // Save the document as a multi‑page TIFF.
        doc.Save(tiffPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Output the result (optional).
        long fileSize = new FileInfo(tiffPath).Length;
        Console.WriteLine($"TIFF saved to '{tiffPath}' ({fileSize} bytes).");
    }
}
