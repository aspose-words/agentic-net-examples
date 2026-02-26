using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcodeDoc
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define barcode parameters – here we generate a QR code.
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

        // Generate the barcode image using the built‑in barcode generator.
        // The GetBarcodeImage method returns a stream containing the image data.
        using (Stream imgStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParams))
        {
            // Insert the generated image into the document at the current cursor position.
            builder.InsertImage(imgStream);
        }

        // Save the document to a DOCX file.
        doc.Save("GeneratedBarcode.docx");
    }
}
