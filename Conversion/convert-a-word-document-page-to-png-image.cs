using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class WordToPngConverter
{
    /// <summary>
    /// Converts a single page of a Word document to a PNG image.
    /// </summary>
    /// <param name="inputFile">Full path to the source .doc/.docx file.</param>
    /// <param name="outputFile">Full path where the PNG image will be saved.</param>
    /// <param name="pageIndex">Zero‑based index of the page to render.</param>
    public static void ConvertPageToPng(string inputFile, string outputFile, int pageIndex)
    {
        // Load the document from the file system.
        Document doc = new Document(inputFile);

        // Create ImageSaveOptions for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Restrict rendering to the specified page.
        options.PageSet = new PageSet(pageIndex);

        // Optional: set resolution (dpi) if higher quality is required.
        // options.Resolution = 300;

        // Save the selected page as a PNG image.
        doc.Save(outputFile, options);
    }
}

public class Program
{
    // Entry point required for a console application.
    public static void Main(string[] args)
    {
        // Expected arguments: <inputFile> <outputFile> <pageIndex>
        if (args.Length != 3)
        {
            Console.WriteLine("Usage: <inputFile> <outputFile> <pageIndex (zero‑based)>");
            return;
        }

        string inputFile = args[0];
        string outputFile = args[1];
        if (!int.TryParse(args[2], out int pageIndex))
        {
            Console.WriteLine("Invalid page index.");
            return;
        }

        try
        {
            WordToPngConverter.ConvertPageToPng(inputFile, outputFile, pageIndex);
            Console.WriteLine($"Page {pageIndex + 1} saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
