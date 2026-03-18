using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Define barcode parameters (QR code in this example).
        BarcodeParameters barcodeParams = new BarcodeParameters
        {
            BarcodeType = "QR",
            BarcodeValue = "ABC123",
            BackgroundColor = "0xF8BD69",
            ForegroundColor = "0xB5413B",
            ErrorCorrectionLevel = "3",
            ScalingFactor = "250",
            SymbolHeight = "1000",
            SymbolRotation = "0"
        };

        // Generate the barcode image as a stream using the built‑in generator.
        using (Stream barcodeStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParams))
        {
            // Copy the stream into a MemoryStream (byte array) so we can reuse it.
            using (MemoryStream imageBytes = new MemoryStream())
            {
                barcodeStream.CopyTo(imageBytes);
                imageBytes.Position = 0; // Reset position before using.

                // Configure watermark appearance.
                ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
                {
                    Scale = 5,          // Scale factor (5 = 5× the original size).
                    IsWashout = false   // Disable washout effect for a solid image.
                };

                // Apply the barcode image as a document watermark.
                doc.Watermark.SetImage(imageBytes, watermarkOptions);
            }
        }

        // Save the resulting document.
        doc.Save("BarcodeWatermark.docx");
    }
}
