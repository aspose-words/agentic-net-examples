using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder to edit it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up barcode parameters for a QR code.
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

        // Generate the barcode image using the built‑in barcode generator.
        using (Stream img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters))
        {
            // Insert the generated image into the document.
            img.Position = 0;
            builder.InsertImage(img);
        }

        // Save the resulting document as a PDF file.
        doc.Save("BarcodeOutput.pdf");
    }
}
