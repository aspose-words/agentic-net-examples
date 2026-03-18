using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ImageToPdfConversion
{
    public static class Converter
    {
        /// <summary>
        /// Converts a PNG image file to a PDF document.
        /// </summary>
        /// <param name="pngFilePath">Full path to the source PNG image.</param>
        /// <param name="pdfFilePath">Full path where the resulting PDF will be saved.</param>
        public static void ConvertPngToPdf(string pngFilePath, string pdfFilePath)
        {
            // Create a new empty document.
            Document doc = new Document();

            // Use DocumentBuilder to insert the PNG image into the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(pngFilePath);

            // Save the document as PDF using the SaveFormat.Pdf enumeration value.
            doc.Save(pdfFilePath, SaveFormat.Pdf);
        }

        // Example usage.
        public static void Main()
        {
            string pngPath = @"C:\Images\sample.png";
            string pdfPath = @"C:\Output\sample.pdf";

            ConvertPngToPdf(pngPath, pdfPath);

            Console.WriteLine("Conversion completed.");
        }
    }
}
