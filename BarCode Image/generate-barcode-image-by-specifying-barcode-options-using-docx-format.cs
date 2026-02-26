using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will generate a QR code.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);

        // Set the barcode type to QR.
        barcodeField.BarcodeType = "QR";

        // Set the value that will be encoded.
        barcodeField.BarcodeValue = "Aspose.Words";

        // Apply additional options (colors, error correction, scaling, size, rotation).
        // All options are string values as defined by the MERGEBARCODE field syntax.
        barcodeField.BackgroundColor = "F8BD69";          // Background color in RRGGBB hex.
        barcodeField.ForegroundColor = "B5413B";          // Foreground color in RRGGBB hex.
        barcodeField.ErrorCorrectionLevel = "3";          // QR error correction level (0‑3).
        barcodeField.ScalingFactor = "250";               // Scale the symbol to 250 %.
        barcodeField.SymbolHeight = "1000";               // Height in TWIPS (1/1440 inch).
        barcodeField.SymbolRotation = "0";                // No rotation.

        // Insert a line break after the field for readability.
        builder.Writeln();

        // Update all fields in the document so the barcode is rendered.
        doc.UpdateFields();

        // Save the document in DOCX format.
        doc.Save("GeneratedBarcode.docx");
    }
}
