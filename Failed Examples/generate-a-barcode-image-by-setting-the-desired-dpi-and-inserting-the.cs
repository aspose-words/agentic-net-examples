// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using System.Drawing;                     // For Image manipulation
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeExample
{
    // Minimal stub implementation of IBarcodeGenerator.
    // In a real scenario replace this with an actual generator that creates barcode images.
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        public Stream GetBarcodeImage(BarcodeParameters parameters)
        {
            // Placeholder: return an empty PNG image stream.
            // Replace with actual barcode generation logic.
            using (var bmp = new Bitmap(200, 200))
            {
                using (var graphics = Graphics.FromImage(bmp))
                {
                    graphics.Clear(Color.White);
                }

                var ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                ms.Position = 0;
                return ms;
            }
        }

        public Stream GetOldBarcodeImage(BarcodeParameters parameters)
        {
            // For simplicity, reuse the same implementation.
            return GetBarcodeImage(parameters);
        }
    }

    class Program
    {
        static void Main()
        {
            // Directory where the resulting document will be saved.
            string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // Create a new blank document and a DocumentBuilder for editing.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Assign the custom barcode generator to the document.
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

            // Generate the barcode image as a stream.
            using (Stream rawImageStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters))
            {
                // Load the image into a System.Drawing.Image to modify its DPI.
                using (Image image = Image.FromStream(rawImageStream))
                {
                    // Desired DPI (dots per inch).
                    const float desiredDpi = 300f;

                    // Set the image resolution.
                    image.SetResolution(desiredDpi, desiredDpi);

                    // Save the modified image back to a memory stream in PNG format.
                    using (MemoryStream dpiAdjustedStream = new MemoryStream())
                    {
                        image.Save(dpiAdjustedStream, System.Drawing.Imaging.ImageFormat.Png);
                        dpiAdjustedStream.Position = 0; // Reset stream position for insertion.

                        // Insert the DPI‑adjusted image into the document.
                        builder.InsertImage(dpiAdjustedStream);
                    }
                }
            }

            // Save the document as DOCX.
            string outputPath = Path.Combine(artifactsDir, "BarcodeWithDPI.docx");
            doc.Save(outputPath);
        }
    }
}
