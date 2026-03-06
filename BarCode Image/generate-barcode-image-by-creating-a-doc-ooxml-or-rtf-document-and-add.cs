using System;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcode
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field.
        // This field will display a QR code with custom colors and scaling.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set barcode properties.
        barcodeField.BarcodeType = "QR";                 // Type of barcode.
        barcodeField.BarcodeValue = "ABC123";            // Data to encode.
        barcodeField.BackgroundColor = "0xF8BD69";       // Background colour (hex).
        barcodeField.ForegroundColor = "0xB5413B";       // Foreground colour (hex).
        barcodeField.ErrorCorrectionLevel = "3";         // QR error correction level.
        barcodeField.ScalingFactor = "250";              // Scale the symbol to 250%.
        barcodeField.SymbolHeight = "1000";              // Height in TWIPS (1/1440 inch).
        barcodeField.SymbolRotation = "0";               // No rotation.

        // Optional: add a line break after the field for readability.
        builder.Writeln();

        // Save the document in DOCX format.
        doc.Save("BarcodeDisplay.docx");
    }
}
