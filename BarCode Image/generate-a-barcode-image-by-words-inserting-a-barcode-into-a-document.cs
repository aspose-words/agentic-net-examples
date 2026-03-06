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
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";                 // Set barcode type.
        barcodeField.BarcodeValue = "ABC123";            // Data to encode.
        barcodeField.BackgroundColor = "0xF8BD69";       // Background color.
        barcodeField.ForegroundColor = "0xB5413B";       // Foreground (bars) color.
        barcodeField.ErrorCorrectionLevel = "3";         // QR error correction level.
        barcodeField.ScalingFactor = "250";              // Scale the symbol.
        barcodeField.SymbolHeight = "1000";              // Height in TWIPS.
        barcodeField.SymbolRotation = "0";               // No rotation.

        // Add a line break after the barcode (optional).
        builder.Writeln();

        // Save the document as a DOCX file.
        doc.Save("BarcodeDocument.docx");
    }
}
