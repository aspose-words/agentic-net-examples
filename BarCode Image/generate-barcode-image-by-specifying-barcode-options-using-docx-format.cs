using System;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcodeDocx
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field. The field will be updated immediately because we pass true.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);

        // Specify barcode options. All properties of FieldMergeBarcode are strings, so numeric values must be supplied as string literals.
        barcodeField.BarcodeType = "QR";                     // QR code type.
        barcodeField.BarcodeValue = "ABC123";                // Data to encode.
        barcodeField.BackgroundColor = "F8BD69";            // Background color (hex RGB without leading '#').
        barcodeField.ForegroundColor = "B5413B";            // Foreground color (hex RGB without leading '#').
        barcodeField.ErrorCorrectionLevel = "3";            // QR error correction level (0‑3).
        barcodeField.ScalingFactor = "250";                 // Scale to 250 %.
        barcodeField.SymbolHeight = "1000";                 // Height in TWIPS (1/1440 inch).
        barcodeField.SymbolRotation = "0";                  // No rotation.

        // Optionally add a line break after the field.
        builder.Writeln();

        // Save the document in DOCX format.
        doc.Save("BarcodeDocument.docx");
    }
}
