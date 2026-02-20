using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ConvertMultiPageDocxToImages
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the DOCX document from disk.
            Document doc = new Document("input.docx");

            // Configure image save options to render each page as a separate PNG image.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render all pages (this is the default, but set explicitly for clarity).
                PageSet = PageSet.All,
                // Optional: set the resolution (dpi) for higher quality images.
                // Resolution = 300,
                // Optional: specify a custom file name pattern.
                // OutputFileName = "output.png"
            };

            // Save the document. Aspose.Words will create one PNG file per page,
            // naming them like "output.png", "output_1.png", "output_2.png", etc.
            doc.Save("output.png", saveOptions);
        }
    }
}
