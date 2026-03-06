// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using System.Drawing;                     // For setting image DPI
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Assign a custom barcode generator (implementation must be provided elsewhere).
        // This generator will be used to create the barcode image.
        doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Define barcode parameters.
        BarcodeParameters parameters = new BarcodeParameters
        {
            BarcodeType = "QR",                 // Type of barcode (e.g., QR, CODE39, EAN13, etc.)
            BarcodeValue = "ABC123",            // Data to encode
            ScalingFactor = "250",              // Scaling factor influences image size
            SymbolHeight = "1000",              // Height in twips (1/1440 inch)
            SymbolRotation = "0"
        };

        // Generate the barcode image as a stream.
        using (Stream rawImageStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(parameters))
        {
            // Load the image into a System.Drawing.Image to modify its DPI.
            using (Image img = Image.FromStream(rawImageStream))
            {
                // Set the desired DPI (e.g., 300 DPI).
                const float desiredDpi = 300f;
                img.SetResolution(desiredDpi, desiredDpi);

                // Save the modified image back to a memory stream.
                using (MemoryStream dpiAdjustedStream = new MemoryStream())
                {
                    // Preserve the original format (e.g., PNG, JPEG) by using the raw stream's format.
                    // Here we default to PNG; adjust as needed.
                    img.Save(dpiAdjustedStream, System.Drawing.Imaging.ImageFormat.Png);
                    dpiAdjustedStream.Position = 0; // Reset position before insertion.

                    // Insert the DPI‑adjusted image into the document.
                    builder.InsertImage(dpiAdjustedStream);
                }
            }
        }

        // Save the document to a DOCX file.
        doc.Save("BarcodeWithDPI.docx");
    }
}

// Placeholder for a user‑implemented barcode generator.
// The actual implementation must conform to IBarcodeGenerator.
public class CustomBarcodeGenerator : IBarcodeGenerator
{
    public Stream GetBarcodeImage(BarcodeParameters parameters)
    {
        // Implementation should generate a barcode image based on the parameters
        // and return it as a Stream. This placeholder throws to indicate it is not implemented.
        throw new NotImplementedException("Custom barcode generation logic must be provided.");
    }

    public Stream GetOldBarcodeImage(BarcodeParameters parameters)
    {
        // Implementation for legacy barcode fields (optional).
        throw new NotImplementedException("Custom barcode generation logic must be provided.");
    }
}
