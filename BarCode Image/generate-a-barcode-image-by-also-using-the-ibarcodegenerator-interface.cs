using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeConversionExample
{
    // Simple implementation of IBarcodeGenerator that returns a tiny placeholder PNG.
    // In a real scenario you would replace this with a proper barcode generation library.
    public class SimpleBarcodeGenerator : IBarcodeGenerator
    {
        // Generates an image for DISPLAYBARCODE fields.
        public Stream GetBarcodeImage(BarcodeParameters parameters)
        {
            return GeneratePlaceholderImage(parameters);
        }

        // Generates an image for old BARCODE fields.
        public Stream GetOldBarcodeImage(BarcodeParameters parameters)
        {
            return GeneratePlaceholderImage(parameters);
        }

        private Stream GeneratePlaceholderImage(BarcodeParameters parameters)
        {
            // A 1x1 pixel transparent PNG (base64 encoded).
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            // Return a MemoryStream that is ready for reading.
            return new MemoryStream(pngBytes);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file that contains barcode fields.
            string inputDocxPath = "Barcodes.docx";

            // Path where the resulting PDF will be saved.
            string outputPdfPath = "Barcodes.pdf";

            // Load the DOCX document.
            Document doc = new Document(inputDocxPath);

            // Assign the custom barcode generator to the document.
            doc.FieldOptions.BarcodeGenerator = new SimpleBarcodeGenerator();

            // Update all fields so that barcode images are generated.
            doc.UpdateFields();

            // Save the document as PDF. All barcode fields are now rendered as images.
            doc.Save(outputPdfPath, SaveFormat.Pdf);
        }
    }
}
