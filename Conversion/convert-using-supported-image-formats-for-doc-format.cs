using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.doc");

        // Define the image formats that are supported for rendering a DOC.
        SaveFormat[] imageFormats = new SaveFormat[]
        {
            SaveFormat.Png,   // Portable Network Graphics
            SaveFormat.Jpeg,  // JPEG image
            SaveFormat.Bmp,   // Bitmap
            SaveFormat.Tiff,  // TIFF image
            SaveFormat.Emf,   // Enhanced Metafile (vector)
            SaveFormat.Eps,   // Encapsulated PostScript (vector)
            SaveFormat.WebP,  // WebP image
            SaveFormat.Svg    // Scalable Vector Graphics
        };

        // Iterate through each page of the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Export the current page to each of the selected image formats.
            foreach (SaveFormat format in imageFormats)
            {
                // Create ImageSaveOptions for the desired format.
                ImageSaveOptions options = new ImageSaveOptions(format)
                {
                    // Render only the current page.
                    PageSet = new PageSet(pageIndex),

                    // Set a reasonable resolution (dpi) for raster formats.
                    Resolution = 300
                };

                // Determine the file extension for the chosen format.
                string extension = FileFormatUtil.SaveFormatToExtension(format);

                // Build the output file name (e.g., Output_Page1.png).
                string outputPath = $"Output_Page{pageIndex + 1}{extension}";

                // Save the page as an image using the configured options.
                doc.Save(outputPath, options);
            }
        }
    }
}
