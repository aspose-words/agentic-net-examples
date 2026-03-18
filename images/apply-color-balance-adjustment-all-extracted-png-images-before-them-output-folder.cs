using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractAndAdjustImages
{
    static void Main()
    {
        // Path to the source Word document.
        string sourceDocPath = @"C:\Input\Document.docx";

        // Folder where the adjusted PNG images will be saved.
        string outputFolder = @"C:\Output\Images";

        // Load the document.
        Document doc = new Document(sourceDocPath);

        // Ensure the output folder exists.
        System.IO.Directory.CreateDirectory(outputFolder);

        // Iterate through each page of the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Create ImageSaveOptions for PNG format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Adjust color balance: brightness and contrast.
                // Values are in the range 0..1 where 0.5 is the default.
                ImageBrightness = 0.6f, // Slightly brighter.
                ImageContrast   = 0.7f, // Slightly higher contrast.

                // Render only the current page.
                PageSet = new PageSet(pageIndex)
            };

            // Build the output file name.
            string outputPath = System.IO.Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the rendered page as a PNG image with the specified adjustments.
            doc.Save(outputPath, options);
        }
    }
}
