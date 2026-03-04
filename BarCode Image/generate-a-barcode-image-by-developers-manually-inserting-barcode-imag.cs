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

        // Prepare barcode parameters for a QR code.
        BarcodeParameters barcodeParams = new BarcodeParameters
        {
            BarcodeType = "QR",                 // Type of barcode.
            BarcodeValue = "ABC123",            // Data to encode.
            ScalingFactor = "300",              // Scale the symbol (percentage).
            SymbolHeight = "2000",              // Height in TWIPS (1/1440 inch).
            SymbolRotation = "0",               // No rotation.
            BackgroundColor = "0xF8BD69",       // Optional background color.
            ForegroundColor = "0xB5413B"        // Optional foreground color.
        };

        // Generate the barcode image using the built‑in barcode generator.
        // The generator returns a stream containing the image data.
        using (Stream barcodeImage = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParams))
        {
            // Insert the image into the document.
            // Width and height are specified in points (1 point = 1/72 inch).
            // Here we set 100 points (~1.39 cm) for both dimensions.
            builder.InsertImage(barcodeImage, 100, 100);
        }

        // Save the document in DOCX format.
        doc.Save("BarcodeDocument.docx");
    }
}
