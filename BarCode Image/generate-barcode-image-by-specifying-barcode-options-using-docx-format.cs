using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field that will render a QR code.
        FieldDisplayBarcode barcode = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set the barcode type and the data to encode.
        barcode.BarcodeType = "QR";
        barcode.BarcodeValue = "ABC123";

        // Apply custom visual options.
        barcode.BackgroundColor = "0xF8BD69";   // Light orange background.
        barcode.ForegroundColor = "0xB5413B";   // Dark red bars.
        barcode.ErrorCorrectionLevel = "3";     // Highest error correction.
        barcode.ScalingFactor = "250";          // 250 % scaling.
        barcode.SymbolHeight = "1000";          // Height in TWIPS (≈0.69 in).
        barcode.SymbolRotation = "0";           // No rotation.

        // Save the document in DOCX format.
        doc.Save("BarcodeDocument.docx");
    }
}
