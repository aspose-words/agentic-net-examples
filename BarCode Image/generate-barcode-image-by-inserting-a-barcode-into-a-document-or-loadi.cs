using System;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will generate a QR code.
        // The field is inserted with the "updateField" flag set to true so we can
        // immediately update it after setting its properties.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);

        // Configure the barcode parameters. All properties are strings, so values must be supplied as strings.
        barcodeField.BarcodeType = "QR";                     // QR code type.
        barcodeField.BarcodeValue = "ABC123";               // Data to encode.
        barcodeField.BackgroundColor = "F8BD69";            // Background colour (hex RGB, without leading '#').
        barcodeField.ForegroundColor = "B5413B";            // Foreground colour (hex RGB).
        barcodeField.ErrorCorrectionLevel = "3";            // Highest error correction (as string).
        barcodeField.ScalingFactor = "250";                 // 250 % scaling (as string).
        barcodeField.SymbolHeight = "1000";                 // Height in TWIPS (as string).
        barcodeField.SymbolRotation = "0";                  // No rotation (as string).

        // Force the field to recalculate and render the barcode image.
        barcodeField.Update();

        // Save the document in DOCX format.
        doc.Save("BarcodeDocument.docx");
    }
}
