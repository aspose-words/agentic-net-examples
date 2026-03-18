using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Folder containing the source DOCX files
        string inputFolder = @"C:\InputDocs";

        // Folder where the resulting TIFF files will be saved
        string outputFolder = @"C:\OutputTiffs";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each DOCX file in the input folder
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document (uses Document(string) constructor)
            Document doc = new Document(docxPath);

            // Create ImageSaveOptions for TIFF output (uses ImageSaveOptions(SaveFormat) constructor)
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.Resolution = 200;                     // Set DPI to 200
            options.TiffCompression = TiffCompression.Ccitt3; // Apply CCITT3 compression

            // Build the output file name with .tiff extension
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(docxPath);
            string tiffPath = Path.Combine(outputFolder, fileNameWithoutExt + ".tiff");

            // Save the document as a TIFF image using the configured options (uses Document.Save(string, SaveOptions))
            doc.Save(tiffPath, options);
        }
    }
}
