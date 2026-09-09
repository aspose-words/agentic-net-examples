using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with more than five pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7) // No break after the last page.
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Prepare the folder for the output file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string tiffPath = Path.Combine(outputDir, "FirstFivePages.tiff");

        // Configure ImageSaveOptions for TIFF and limit to the first five pages.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // Page indices are zero‑based, so pages 0‑4 correspond to the first five pages.
        options.PageSet = new PageSet(0, 1, 2, 3, 4);
        // Use the default multi‑page layout for TIFF (each page as a separate frame).
        options.PageLayout = MultiPageLayout.TiffFrames();

        // Save the document as a multi‑page TIFF containing only the first five pages.
        doc.Save(tiffPath, options);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("The TIFF file was not created.", tiffPath);

        Console.WriteLine($"TIFF saved successfully to: {tiffPath}");
    }
}
