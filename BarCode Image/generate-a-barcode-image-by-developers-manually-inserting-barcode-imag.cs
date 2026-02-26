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

        // -----------------------------------------------------------------
        // Insert a QR code barcode with custom colors and a specific size.
        // -----------------------------------------------------------------
        BarcodeParameters qrParams = new BarcodeParameters
        {
            BarcodeType = "QR",
            BarcodeValue = "ABC123",
            BackgroundColor = "0xF8BD69",
            ForegroundColor = "0xB5413B",
            ErrorCorrectionLevel = "3",
            ScalingFactor = "250",
            SymbolHeight = "1000", // height in twips (1/1440 inch)
            SymbolRotation = "0"
        };

        using (Stream img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(qrParams))
        {
            // Insert the barcode image inline, scaling it to 200 pt × 200 pt.
            // 1 point = 1/72 inch.
            builder.InsertImage(img, 200, 200);
        }

        // -----------------------------------------------------------------
        // Insert a CODE39 barcode that includes start/stop characters.
        // -----------------------------------------------------------------
        builder.Writeln(); // Move to a new line.

        BarcodeParameters code39Params = new BarcodeParameters
        {
            BarcodeType = "CODE39",
            BarcodeValue = "12345ABCDE",
            AddStartStopChar = true
        };

        using (Stream img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(code39Params))
        {
            // Insert the barcode image inline, scaling it to 150 pt × 50 pt.
            builder.InsertImage(img, 150, 50);
        }

        // -----------------------------------------------------------------
        // Save the resulting DOCX file.
        // -----------------------------------------------------------------
        doc.Save("Barcodes.docx");
    }
}
