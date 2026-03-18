using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PngToGrayscaleBmp
{
    static void Main()
    {
        // Folder containing the extracted PNG images.
        string sourceFolder = @"C:\SourcePngs";

        // Secure folder where the grayscale BMP files will be archived.
        string archiveFolder = @"C:\SecureArchive";

        // Ensure the archive folder exists.
        Directory.CreateDirectory(archiveFolder);

        // Process each PNG file in the source folder.
        foreach (string pngPath in Directory.GetFiles(sourceFolder, "*.png"))
        {
            // ---------- Create ----------
            // Create a new empty Word document.
            Document doc = new Document();

            // Create a DocumentBuilder to work with the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // ---------- Load ----------
            // Insert the PNG image into the document.
            builder.InsertImage(pngPath);

            // ---------- Save ----------
            // Configure save options to render the page as a BMP image
            // and force the output to be grayscale.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Bmp)
            {
                ImageColorMode = ImageColorMode.Grayscale
            };

            // Build the destination BMP file name.
            string bmpFileName = Path.GetFileNameWithoutExtension(pngPath) + ".bmp";
            string bmpPath = Path.Combine(archiveFolder, bmpFileName);

            // Save the rendered page (containing the grayscale image) as BMP.
            doc.Save(bmpPath, saveOptions);
        }
    }
}
