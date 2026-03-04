using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcode
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define barcode parameters (QR code example).
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

        // Generate the barcode image using the document's barcode generator.
        using (Stream imgStream = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(parameters))
        {
            // Insert the generated image into the document.
            builder.InsertImage(imgStream);
        }

        // Save the document in DOCX format.
        doc.Save("BarcodeQR.docx");
    }
}
