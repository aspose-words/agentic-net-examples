using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class JpegToTiffConverter
{
    static void Main()
    {
        // Folder containing the source JPEG images.
        string jpegFolder = @"C:\Images\Jpeg";
        // Destination folder for the converted TIFF images.
        string tiffFolder = @"C:\Images\Tiff";

        // Ensure the destination folder exists.
        Directory.CreateDirectory(tiffFolder);

        // Process each JPEG file in the source folder.
        foreach (string jpegPath in Directory.GetFiles(jpegFolder, "*.jpg"))
        {
            // Load the JPEG image into a new Word document by inserting it.
            Document doc = new Document();                                   // create document
            DocumentBuilder builder = new DocumentBuilder(doc);               // create builder
            builder.InsertImage(jpegPath);                                   // insert JPEG

            // Configure high‑resolution TIFF output with LZW compression.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff); // create ImageSaveOptions
            tiffOptions.TiffCompression = TiffCompression.Lzw;               // set LZW compression
            tiffOptions.Resolution = 300;                                     // high resolution (300 DPI)

            // Build the output file name with .tiff extension.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(jpegPath);
            string tiffPath = Path.Combine(tiffFolder, fileNameWithoutExt + ".tiff");

            // Save the document as a TIFF image using the configured options.
            doc.Save(tiffPath, tiffOptions);                                 // save document
        }
    }
}
