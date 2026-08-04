using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths for the sample DOCX and the resulting TIFF.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);
        string docPath = Path.Combine(dataDir, "Sample.docx");
        string tiffPath = Path.Combine(dataDir, "Multipage.tiff");

        // Create a sample DOCX with several pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the DOCX (optional, just to demonstrate creation of a source file).
        doc.Save(docPath);

        // Prepare ImageSaveOptions to render a multipage TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use the TiffFrames layout so each page becomes a separate frame in the TIFF.
            PageLayout = MultiPageLayout.TiffFrames()
        };

        // Render the document to a multipage TIFF.
        doc.Save(tiffPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("The multipage TIFF was not created.", tiffPath);

        // Output the result path (no interactive prompts).
        Console.WriteLine($"Multipage TIFF created at: {tiffPath}");
    }
}
