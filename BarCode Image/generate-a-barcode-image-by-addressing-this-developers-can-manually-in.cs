using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set up barcode parameters (QR code example).
        BarcodeParameters parameters = new BarcodeParameters
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

        // Generate the barcode image using the built‑in generator.
        using (Stream imgStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(parameters))
        {
            // Insert the generated image into the document.
            builder.InsertImage(imgStream);
        }

        // Save the document to a DOCX file.
        doc.Save("BarcodeDocument.docx");
    }
}
