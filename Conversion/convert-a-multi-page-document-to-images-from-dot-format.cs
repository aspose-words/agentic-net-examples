using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class DotToImagesConverter
{
    /// <summary>
    /// Loads a multi‑page DOT (Word template) file and converts each page to an image.
    /// Returns an array of streams, each containing one page image in PNG format.
    /// </summary>
    /// <param name="dotFilePath">Full path to the .dot file.</param>
    /// <returns>Array of memory streams, each holding a PNG image of a page.</returns>
    public static Stream[] ConvertDotToImages(string dotFilePath)
    {
        // Load the DOT document using the Document constructor (lifecycle rule).
        Document doc = new Document(dotFilePath);

        // Prepare a list to hold the image streams.
        List<Stream> imageStreams = new List<Stream>();

        // Create an ImageSaveOptions object for PNG output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Iterate over all pages in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Set the PageSet to render only the current page (zero‑based index).
            options.PageSet = new PageSet(pageIndex);

            // Use a MemoryStream to capture the image data (save rule).
            MemoryStream pageStream = new MemoryStream();

            // Save the selected page to the stream using the ImageSaveOptions.
            doc.Save(pageStream, options);

            // Reset the stream position so it can be read by the caller.
            pageStream.Position = 0;

            // Add the stream to the collection.
            imageStreams.Add(pageStream);
        }

        // Return the collection as an array.
        return imageStreams.ToArray();
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Simple argument validation.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: dotnet run <path-to-dot-file> [output-folder]");
            return;
        }

        string dotPath = args[0];
        string outputFolder = args.Length > 1 ? args[1] : Path.Combine(Directory.GetCurrentDirectory(), "output");

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Convert the DOT file to per‑page PNG streams.
        Stream[] pageImages = DotToImagesConverter.ConvertDotToImages(dotPath);

        // Write each stream to a physical PNG file.
        for (int i = 0; i < pageImages.Length; i++)
        {
            string outFile = Path.Combine(outputFolder, $"Page_{i + 1}.png");
            using (FileStream file = new FileStream(outFile, FileMode.Create, FileAccess.Write))
            {
                pageImages[i].CopyTo(file);
            }
            Console.WriteLine($"Saved page {i + 1} to {outFile}");
        }
    }
}
