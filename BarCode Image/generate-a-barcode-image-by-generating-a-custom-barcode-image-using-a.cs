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

        // Insert a MERGEBARCODE field that will generate a QR code.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);

        // Set the barcode type and the value to encode.
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "ABC123";

        // Optional: customize colors and other appearance settings.
        // All these properties are strings, so we provide the values as string literals.
        barcodeField.BackgroundColor = "F8BD69";      // Background color (hex RGB without '#').
        barcodeField.ForegroundColor = "B5413B";      // Foreground (bars) color.
        barcodeField.ErrorCorrectionLevel = "3";      // QR error correction level (0‑3).
        barcodeField.ScalingFactor = "250";           // Scale the symbol (percentage).
        barcodeField.SymbolHeight = "1000";           // Height in TWIPS (1/1440 inch).
        barcodeField.SymbolRotation = "0";            // No rotation.

        // Force the field to update so the barcode image is generated.
        barcodeField.Update();

        // Save the document containing the generated barcode.
        doc.Save("BarcodeDocument.docx");
    }
}
