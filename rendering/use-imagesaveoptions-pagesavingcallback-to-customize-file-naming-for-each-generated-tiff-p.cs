using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder where the TIFF pages will be saved.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        // Configure ImageSaveOptions for TIFF and assign a callback to name each page file.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.PageSavingCallback = new CustomPageSavingCallback(outputDir);
        saveOptions.Resolution = 300; // DPI (optional)

        // Save the document; the callback will create separate TIFF files per page.
        // The file name passed here is ignored because the callback provides its own names.
        doc.Save(Path.Combine(outputDir, "placeholder.tiff"), saveOptions);

        // Verify that the expected number of TIFF files were created.
        string[] tiffFiles = Directory.GetFiles(outputDir, "Page_*.tiff");
        if (tiffFiles.Length != doc.PageCount)
            throw new InvalidOperationException($"Expected {doc.PageCount} TIFF files, but found {tiffFiles.Length}.");

        // List the generated files.
        foreach (string file in tiffFiles)
            Console.WriteLine($"Created: {file}");
    }

    // Callback that sets a custom file name for each page saved as a TIFF image.
    private class CustomPageSavingCallback : IPageSavingCallback
    {
        private readonly string _folder;

        public CustomPageSavingCallback(string folder)
        {
            _folder = folder;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // Create a file name like "Page_1.tiff", "Page_2.tiff", etc.
            string fileName = Path.Combine(_folder, $"Page_{args.PageIndex + 1}.tiff");

            // Use the file name directly.
            args.PageFileName = fileName;

            // Alternatively, provide a stream (shown here for completeness).
            args.PageStream = new FileStream(fileName, FileMode.Create);

            // Ensure Aspose.Words closes the stream after writing.
            args.KeepPageStreamOpen = false;
        }
    }
}
