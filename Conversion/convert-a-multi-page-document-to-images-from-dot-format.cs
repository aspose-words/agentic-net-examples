using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class DotToImagesConverter
{
    public static void Main()
    {
        // Path to the source DOT file (Word template)
        const string inputFilePath = "Template.dot";

        // Directory where the resulting page images will be saved
        const string outputDirectory = "PageImages";
        Directory.CreateDirectory(outputDirectory);

        // Load the DOT document. No LoadOptions are required for a plain template.
        Document doc = new Document(inputFilePath);

        // Prepare image save options – PNG format with 300 DPI resolution.
        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            Resolution = 300
        };

        // Export each page of the document to a separate PNG file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Tell Aspose.Words which page to render.
            imageSaveOptions.PageSet = new PageSet(pageIndex);

            string outputPath = Path.Combine(outputDirectory, $"Page_{pageIndex + 1}.png");
            doc.Save(outputPath, imageSaveOptions);
        }
    }
}
