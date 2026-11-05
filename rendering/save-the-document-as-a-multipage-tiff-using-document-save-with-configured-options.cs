using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with multiple pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure TIFF save options for a multipage output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.PageSet = PageSet.All; // Save all pages.

        // Save the document as a multipage TIFF file.
        string outputPath = "output.tiff";
        doc.Save(outputPath, tiffOptions);

        // Verify that the TIFF file was created and is not empty.
        if (!File.Exists(outputPath))
            throw new Exception("The TIFF file was not created.");

        FileInfo fileInfo = new FileInfo(outputPath);
        if (fileInfo.Length == 0)
            throw new Exception("The TIFF file is empty.");
    }
}
