using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractAndAdjustImages
{
    static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Folder where the adjusted PNG images will be saved.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");
        Directory.CreateDirectory(outputFolder);

        // Iterate through each page of the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                ImageBrightness = 0.6f, // Slightly brighter.
                ImageContrast   = 0.7f, // Slightly higher contrast.
                PageSet = new PageSet(pageIndex) // Render only the current page.
            };

            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");
            doc.Save(outputPath, options);
        }

        Console.WriteLine($"Saved {doc.PageCount} page(s) to \"{outputFolder}\".");
    }
}
