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

        // Insert a MERGEBARCODE field that will render a QR code.
        FieldMergeBarcode qrField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        qrField.BarcodeType = "QR";                         // QR code type.
        qrField.BarcodeValue = "https://example.com";       // Data to encode.
        qrField.BackgroundColor = "F8BD69";                // Light orange background (hex string).
        qrField.ForegroundColor = "B5413B";                // Dark red foreground (hex string).
        qrField.ErrorCorrectionLevel = "3";                // Highest error correction (string).
        qrField.ScalingFactor = "250";                     // 250 % scaling (string).
        qrField.SymbolHeight = "1000";                     // Height in TWIPS (string).
        qrField.SymbolRotation = "0";                      // No rotation (string).

        builder.Writeln(); // Add a line break after the barcode.

        // Save the document containing the QR code.
        doc.Save("BarcodeQR.docx");

        // -----------------------------------------------------------------
        // Load the previously saved document and add a CODE39 barcode.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document("BarcodeQR.docx");
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

        loadedBuilder.Writeln("Additional barcode:");

        // Insert a MERGEBARCODE field for a CODE39 barcode.
        FieldMergeBarcode code39Field = (FieldMergeBarcode)loadedBuilder.InsertField(FieldType.FieldMergeBarcode, true);
        code39Field.BarcodeType = "CODE39";          // CODE39 type.
        code39Field.BarcodeValue = "12345ABCDE";    // Data to encode.
        code39Field.AddStartStopChar = true;        // Show start/stop characters.

        // Save the updated document.
        loadedDoc.Save("BarcodeQR_WithCode39.docx");
    }
}
