using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class DotToImageConverter
{
    /// <summary>
    /// Converts each page of a multi‑page DOT (Word template) document to separate image files.
    /// </summary>
    /// <param name="dotFilePath">Full path to the input .dot file.</param>
    /// <param name="outputFolder">Folder where the image files will be saved.</param>
    public void Convert(string dotFilePath, string outputFolder)
    {
        // Ensure the output directory exists.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // Load the DOT document. The Document constructor handles format detection.
        Document doc = new Document(dotFilePath);

        // Iterate through all pages in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Prepare image save options for the current page.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page (zero‑based index).
                PageSet = new PageSet(pageIndex)
            };

            // Build the output file name (e.g., "Page_1.png").
            string outFile = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the rendered page as an image file.
            doc.Save(outFile, imageOptions);
        }
    }
}

public static class Program
{
    /// <summary>
    /// Entry point required for console execution.
    /// </summary>
    public static void Main(string[] args)
    {
        // Simple argument handling – if not provided, use example paths.
        string dotPath = args.Length > 0 ? args[0] : @"C:\Docs\Template.dot";
        string outFolder = args.Length > 1 ? args[1] : @"C:\Docs\Images";

        try
        {
            var converter = new DotToImageConverter();
            converter.Convert(dotPath, outFolder);
            Console.WriteLine($"Conversion completed. Images saved to '{outFolder}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
