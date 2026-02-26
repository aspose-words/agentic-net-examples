using System;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcodesWithFields
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will display a QR code.
        // The second argument (true) tells the builder to update the field result immediately.
        FieldMergeBarcode qrField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        qrField.BarcodeType = "QR";                 // Set the barcode type.
        qrField.BarcodeValue = "Aspose.Words";      // Data to encode.
        qrField.BackgroundColor = "0xF8BD69";       // Optional: background colour.
        qrField.ForegroundColor = "0xB5413B";       // Optional: foreground colour.
        qrField.ErrorCorrectionLevel = "3";         // Optional: QR error correction.
        qrField.ScalingFactor = "250";              // Optional: scaling factor (percentage).
        qrField.SymbolHeight = "1000";              // Optional: height in TWIPS.
        qrField.SymbolRotation = "0";               // Optional: rotation.

        // Add a line break after the QR code.
        builder.Writeln();

        // Insert a MERGEBARCODE field that will display an EAN13 barcode.
        FieldMergeBarcode eanField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        eanField.BarcodeType = "EAN13";
        eanField.BarcodeValue = "501234567890";
        eanField.DisplayText = true;               // Show the numeric value under the bars.
        eanField.PosCodeStyle = "CASE";
        eanField.FixCheckDigit = true;

        // Add a line break after the EAN13 barcode.
        builder.Writeln();

        // Insert a MERGEBARCODE field that will display a CODE39 barcode with start/stop characters.
        FieldMergeBarcode code39Field = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "12345ABCDE";
        code39Field.AddStartStopChar = true;        // Include start/stop characters.

        // Save the document in DOCX format.
        doc.Save("BarcodesWithFields.docx");
    }
}
