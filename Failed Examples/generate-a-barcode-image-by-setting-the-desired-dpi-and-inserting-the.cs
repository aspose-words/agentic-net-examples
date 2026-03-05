// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up barcode parameters (QR code in this example).
        BarcodeParameters barcodeParams = new BarcodeParameters
        {
            BarcodeType = "QR",
            BarcodeValue = "ABC123",
            // Optional: adjust scaling factor if needed.
            ScalingFactor = "250"
        };

        // Generate the barcode image as a stream.
        using (Stream barcodeStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParams))
        {
            // Ensure the stream is positioned at the beginning.
            barcodeStream.Position = 0;

            // Load the image from the stream so we can modify its DPI.
            using (Image img = Image.FromStream(barcodeStream))
            {
                // Set the desired DPI (e.g., 300 DPI for both horizontal and vertical).
                img.SetResolution(300f, 300f);

                // Insert the image with the new DPI into the document.
                builder.InsertImage(img);
            }
        }

        // Save the document as a DOCX file.
        doc.Save("BarcodeDocument.docx");
    }
}
