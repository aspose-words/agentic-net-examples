using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PngToJpegConverter
{
    static void Main()
    {
        // Folder containing the source PNG images.
        string sourceFolder = @"C:\Images\png";

        // Folder where the converted JPEG images will be saved.
        string destinationFolder = @"C:\Images\jpeg";

        // Ensure the destination folder exists.
        Directory.CreateDirectory(destinationFolder);

        // Process each PNG file in the source folder.
        foreach (string pngPath in Directory.GetFiles(sourceFolder, "*.png"))
        {
            // Load the PNG into a new Word document by inserting it as a shape.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(pngPath);

            // Configure image save options for JPEG output.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Preserve original dimensions: keep default resolution (96 DPI) and no scaling.
                Resolution = 96,   // Sets both horizontal and vertical DPI.
                Scale = 1f,        // 100% scale.
                JpegQuality = 100 // Maximum quality (no additional compression).
            };

            // Build the output file name with a .jpg extension.
            string jpegFileName = Path.GetFileNameWithoutExtension(pngPath) + ".jpg";
            string jpegPath = Path.Combine(destinationFolder, jpegFileName);

            // Save the document page (which contains only the image) as a JPEG file.
            doc.Save(jpegPath, jpegOptions);
        }
    }
}
