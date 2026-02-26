// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using System.Drawing; // For Image manipulation
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcodeWithDpi
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Assign a custom barcode generator (implementation must be provided elsewhere).
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Set up barcode parameters (example: QR code).
        BarcodeParameters barcodeParameters = new BarcodeParameters
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

        // Generate the barcode image as a stream.
        using (Stream barcodeStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters))
        {
            // Load the stream into a System.Drawing.Image to modify its DPI.
            using (Image barcodeImage = Image.FromStream(barcodeStream))
            {
                // Desired DPI.
                const float desiredDpi = 300f;

                // Set the image resolution.
                barcodeImage.SetResolution(desiredDpi, desiredDpi);

                // Save the modified image back to a memory stream.
                using (MemoryStream dpiAdjustedStream = new MemoryStream())
                {
                    // Preserve the original format (e.g., PNG, JPEG). Here we use PNG.
                    barcodeImage.Save(dpiAdjustedStream, System.Drawing.Imaging.ImageFormat.Png);
                    dpiAdjustedStream.Position = 0; // Reset stream position for insertion.

                    // Insert the DPI‑adjusted image into the document.
                    builder.InsertImage(dpiAdjustedStream);
                }
            }
        }

        // Save the document as DOCX.
        doc.Save("BarcodeWithDpi.docx");
    }
}

// Placeholder for a custom barcode generator implementation.
// The actual implementation should generate barcode images based on the provided parameters.
public class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(BarcodeParameters parameters)
    {
        // Implementation omitted – return a stream containing a barcode image.
        throw new NotImplementedException();
    }

    public Stream GetOldBarcodeImage(BarcodeParameters parameters)
    {
        // Implementation omitted – return a stream containing a barcode image for old fields.
        throw new NotImplementedException();
    }
}
