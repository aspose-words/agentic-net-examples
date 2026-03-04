using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeExample
{
    // Simple custom barcode generator that returns a placeholder PNG image.
    // The placeholder is a 1x1 pixel transparent PNG encoded as a Base64 string.
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        // Base64 representation of a minimal PNG (1x1 transparent pixel).
        private static readonly byte[] PlaceholderPng = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");

        // Generates a barcode image based on the supplied parameters.
        // For demonstration we ignore the parameters and return the placeholder image.
        public Stream GetBarcodeImage(BarcodeParameters parameters)
        {
            // Return a new MemoryStream each call so the caller can safely dispose it.
            return new MemoryStream(PlaceholderPng, writable: false);
        }

        // Generates an image for the old‑fashioned BARCODE field.
        // Reuse the same placeholder implementation.
        public Stream GetOldBarcodeImage(BarcodeParameters parameters)
        {
            return GetBarcodeImage(parameters);
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Assign the custom barcode generator to the document's field options.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Define barcode parameters (example: QR code).
            BarcodeParameters barcodeParameters = new BarcodeParameters
            {
                BarcodeType = "QR",
                BarcodeValue = "ABC123",
                BackgroundColor = "0xFFFFFF",
                ForegroundColor = "0x000000",
                ErrorCorrectionLevel = "3",
                ScalingFactor = "250",
                SymbolHeight = "1000",
                SymbolRotation = "0"
            };

            // Generate the barcode image and insert it into the document.
            using (Stream imgStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters))
            {
                // Insert the image at the current cursor position.
                builder.InsertImage(imgStream);
            }

            // Save the document as DOCX.
            doc.Save("BarcodeDocument.docx");
        }
    }
}
