using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template document.
        Document doc = new Document("Template.pdf");

        // Iterate through all pages of the document.
        for (int i = 0; i < doc.PageCount; i++)
        {
            // Create ImageSaveOptions for PNG format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(i),

                // Optional: set resolution (dpi) for higher quality.
                Resolution = 300
            };

            // Save the current page as a PNG image.
            string outputPath = $"Page_{i + 1}.png";
            doc.Save(outputPath, options);
        }
    }
}
