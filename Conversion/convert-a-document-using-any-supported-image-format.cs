using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentConversionExample
{
    public class ImageConverter
    {
        /// <summary>
        /// Converts a Word document to an image format supported by Aspose.Words.
        /// </summary>
        /// <param name="inputFilePath">Full path to the source document (e.g., .docx).</param>
        /// <param name="outputFilePath">Full path where the image will be saved (including file name).</param>
        /// <param name="imageFormat">Desired image format (e.g., SaveFormat.Png, SaveFormat.Jpeg, SaveFormat.Tiff, etc.).</param>
        public static void ConvertToImage(string inputFilePath, string outputFilePath, SaveFormat imageFormat)
        {
            // Load the source document from the file system.
            Document doc = new Document(inputFilePath);

            // Configure image save options.
            ImageSaveOptions saveOptions = new ImageSaveOptions(imageFormat)
            {
                // Example: set resolution to 300 DPI for higher quality.
                HorizontalResolution = 300,
                VerticalResolution = 300,
                // Optional: set pixel format, color mode, etc., as needed.
                // PixelFormat = ImagePixelFormat.Format32BppArgb;
                // ImageColorMode = ImageColorMode.None;
            };

            // Save the document as an image. For multi‑page documents, each page will be saved
            // as a separate image file with a numeric suffix (e.g., "output_1.png").
            doc.Save(outputFilePath, saveOptions);
        }

        // Example usage.
        public static void Main()
        {
            string inputPath = @"C:\Docs\SampleDocument.docx";
            string outputPath = @"C:\Docs\SampleDocument.png"; // Base name; Aspose will add page numbers if needed.
            ConvertToImage(inputPath, outputPath, SaveFormat.Png);

            Console.WriteLine("Conversion completed.");
        }
    }
}
