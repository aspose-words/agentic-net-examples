using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeInsertion
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define barcode parameters (QR code in this example).
        // ScalingFactor and SymbolHeight control the resolution/size of the generated image.
        BarcodeParameters barcodeParameters = new BarcodeParameters
        {
            BarcodeType = "QR",
            BarcodeValue = "ABC123",
            ScalingFactor = "300",   // 300% scaling.
            SymbolHeight = "1500"    // Height in twips (1/1440 inch).
        };

        // Generate the barcode image using the built‑in barcode generator.
        using (Stream barcodeImage = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters))
        {
            // Insert the generated image into the document at the current cursor position.
            builder.InsertImage(barcodeImage);
        }

        // Save the document containing the barcode image.
        doc.Save("BarcodeDocument.docx");
    }
}
