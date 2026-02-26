using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SupportedImageFormatsForDoc
{
    static void Main()
    {
        // Load an existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Define the image formats that are supported when rendering a DOC to images.
        SaveFormat[] imageFormats = new SaveFormat[]
        {
            SaveFormat.Png,   // Portable Network Graphics
            SaveFormat.Jpeg,  // Joint Photographic Experts Group
            SaveFormat.Bmp,   // Bitmap
            SaveFormat.Gif,   // Graphics Interchange Format
            SaveFormat.Tiff,  // Tagged Image File Format
            SaveFormat.Emf,   // Enhanced Metafile
            SaveFormat.Eps,   // Encapsulated PostScript
            SaveFormat.WebP,  // WebP
            SaveFormat.Svg    // Scalable Vector Graphics
        };

        // Iterate through each supported image format and save the first page of the document.
        for (int i = 0; i < imageFormats.Length; i++)
        {
            // Create ImageSaveOptions with the desired format.
            ImageSaveOptions options = new ImageSaveOptions(imageFormats[i]);

            // Optionally, set the page to render (e.g., first page).
            options.PageSet = new PageSet(0);

            // Build the output file name based on the format.
            string extension = FileFormatUtil.SaveFormatToExtension(imageFormats[i]);
            string outputPath = $"OutputDocument_Page1{extension}";

            // Save the document (rendered page) to the image file.
            doc.Save(outputPath, options);
        }
    }
}
