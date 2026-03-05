using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will render a QR code.
        FieldMergeBarcode qrField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        qrField.BarcodeType = "QR";                     // QR code type.
        qrField.BarcodeValue = "https://example.com";   // Data to encode.
        qrField.BackgroundColor = "0xF8BD69";           // Light orange background.
        qrField.ForegroundColor = "0xB5413B";           // Dark red foreground.
        qrField.ErrorCorrectionLevel = "3";             // Highest error correction.
        qrField.ScalingFactor = "250";                  // 250 % scaling.
        qrField.SymbolHeight = "1000";                  // Height in TWIPS.
        qrField.SymbolRotation = "0";                   // No rotation.

        builder.Writeln(); // Add a paragraph break.

        // Insert a MERGEBARCODE field that will render a CODE39 barcode.
        FieldMergeBarcode code39Field = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        code39Field.BarcodeType = "CODE39";             // CODE39 type.
        code39Field.BarcodeValue = "12345ABCDE";        // Data to encode.
        code39Field.AddStartStopChar = true;            // Include start/stop characters.

        // Force field calculation so the barcode images appear in the saved document.
        doc.UpdateFields();

        // Save the document in DOCX format.
        doc.Save("Barcodes.docx");
    }
}
