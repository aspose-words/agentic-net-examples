using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeInsertionExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // OPTIONAL: assign a custom barcode generator if you have one.
        // doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

        // Define barcode parameters – here we generate a QR code with a specific resolution.
        BarcodeParameters barcodeParameters = new BarcodeParameters
        {
            BarcodeType = "QR",                 // Type of barcode.
            BarcodeValue = "ABC123",            // Data to encode.
            BackgroundColor = "0xF8BD69",       // Background color (hex).
            ForegroundColor = "0xB5413B",       // Foreground color (hex).
            ErrorCorrectionLevel = "3",         // QR error correction level (0‑3).
            ScalingFactor = "250",              // Scaling factor (percentage) – controls resolution.
            SymbolHeight = "1000",              // Height in twips (1/1440 inch).
            SymbolRotation = "0"                // Rotation (0‑3).
        };

        // Generate the barcode image using the document's barcode generator.
        // The GetBarcodeImage method returns a stream containing the image data.
        using (Stream barcodeImage = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters))
        {
            // Insert the generated image into the document at the current cursor position.
            builder.InsertImage(barcodeImage);

            // If you also need to save the image as a separate file, reset the stream position first.
            barcodeImage.Position = 0;
            using (FileStream file = new FileStream("GeneratedBarcode.png", FileMode.Create))
            {
                barcodeImage.CopyTo(file);
            }
        }

        // Save the document with the inserted barcode to a DOCX file.
        doc.Save("DocumentWithBarcode.docx");
    }
}
