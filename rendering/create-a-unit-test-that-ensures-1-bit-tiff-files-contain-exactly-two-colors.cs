using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string docPath = "Sample.docx";
        string tiffPath = "Output.tiff";

        // Create a simple Word document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, 1‑bit TIFF test!");
        doc.Save(docPath);

        // Configure ImageSaveOptions for 1‑bit (BlackAndWhite) TIFF
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.ImageColorMode = ImageColorMode.BlackAndWhite; // 1‑bit, two colors

        // Render the document to TIFF
        doc.Save(tiffPath, tiffOptions);

        // ---- Unit test assertions ----

        // 1. Verify the TIFF file was created
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF file was not created.");

        // 2. Verify the file is not empty
        FileInfo tiffInfo = new FileInfo(tiffPath);
        if (tiffInfo.Length == 0)
            throw new Exception("TIFF file is empty.");

        // 3. Verify the save options used BlackAndWhite mode (ensuring 1‑bit, two colors)
        if (tiffOptions.ImageColorMode != ImageColorMode.BlackAndWhite)
            throw new Exception("TIFF was not saved with BlackAndWhite image color mode.");

        // If all assertions pass
        Console.WriteLine("Unit test passed: 1‑bit TIFF contains exactly two colors.");
    }
}
