using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string tiffPath = Path.Combine(outputDir, "Ligatures.tiff");

        // Create a new document and add text that contains discretionary ligatures (e.g., "fi", "fl").
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that typically supports ligatures.
        builder.Font.Name = "Calibri";
        builder.Font.Size = 48;
        builder.Writeln("Office");
        builder.Writeln("Affiliation");
        builder.Writeln("Fluff");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // High‑quality rendering improves the chances that ligatures are rendered correctly.
            UseAntiAliasing = true,
            UseHighQualityRendering = true,
            // Optional: increase resolution for clearer output.
            Resolution = 300
        };

        // Render the document to a TIFF file.
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created and has a non‑zero size.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        long fileSize = new FileInfo(tiffPath).Length;
        if (fileSize == 0)
            throw new InvalidOperationException("TIFF file is empty.");

        Console.WriteLine($"TIFF file created successfully at: {tiffPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}
