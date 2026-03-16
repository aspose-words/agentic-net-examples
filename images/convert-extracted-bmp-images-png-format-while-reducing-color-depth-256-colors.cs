using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BmpToPngConverter
{
    static void Main()
    {
        // Folder that contains the source BMP images.
        string sourceFolder = @"C:\Images\Bmp";
        // Folder where the converted PNG images will be saved.
        string targetFolder = @"C:\Images\Png";

        // Ensure the target folder exists.
        Directory.CreateDirectory(targetFolder);

        // Process each BMP file in the source folder.
        foreach (string bmpPath in Directory.GetFiles(sourceFolder, "*.bmp"))
        {
            // Create a new empty Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the BMP image into the document.
            builder.InsertImage(bmpPath);

            // Prepare image save options:
            // - Save format: PNG
            // - (Optional) Reduce colour depth to 256 colours – not directly supported in older
            //   Aspose.Words versions, so we omit the PixelFormat setting. The PNG will be saved
            //   with the default colour depth (24‑bit). If a strict 8‑bit palette is required,
            //   post‑process the file with System.Drawing or ImageMagick.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);
            // saveOptions.PixelFormat = ImagePixelFormat.Format8bppIndexed; // <-- removed – not available in this version

            // Build the output file name with a .png extension.
            string fileName = Path.GetFileNameWithoutExtension(bmpPath) + ".png";
            string pngPath = Path.Combine(targetFolder, fileName);

            // Save the document page (which contains the image) as a PNG file.
            doc.Save(pngPath, saveOptions);
        }
    }
}
