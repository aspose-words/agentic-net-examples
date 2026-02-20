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

        // Insert a MERGEBARCODE field (the field that supports all barcode options).
        // The second argument (true) inserts the field result placeholder.
        builder.InsertField(FieldType.FieldMergeBarcode, true);

        // The newly inserted field is the last field in the document.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)doc.Range.Fields[doc.Range.Fields.Count - 1];

        // Set barcode options – all properties are strings according to the API.
        barcodeField.BarcodeType = "QR";                     // QR code.
        barcodeField.BarcodeValue = "ABC123";               // Data to encode.
        barcodeField.BackgroundColor = "0xF8BD69";          // Background color (hex RGB as string).
        barcodeField.ForegroundColor = "0xB5413B";          // Foreground color (hex RGB as string).
        barcodeField.ErrorCorrectionLevel = "3";           // QR error correction level (0‑3) as string.
        barcodeField.ScalingFactor = "250";                // Scaling factor (percentage) as string.
        barcodeField.SymbolHeight = "1000";                // Height in TWIPS as string.
        barcodeField.SymbolRotation = "0";                 // Rotation (0‑3) as string.

        // Update fields to generate the barcode image.
        doc.UpdateFields();

        // Save the document with the barcode.
        doc.Save("BarcodeDocument.docx");

        // -----------------------------------------------------------------
        // Load the saved document and add another barcode (CODE39) to it.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document("BarcodeDocument.docx");
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

        // Move the cursor to the end of the document.
        loadedBuilder.MoveToDocumentEnd();

        // Insert a paragraph break before the new barcode.
        loadedBuilder.Writeln();

        // Insert a MERGEBARCODE field for CODE39.
        loadedBuilder.InsertField(FieldType.FieldMergeBarcode, true);
        FieldMergeBarcode code39Field = (FieldMergeBarcode)loadedDoc.Range.Fields[loadedDoc.Range.Fields.Count - 1];

        // Set CODE39 barcode options.
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "12345ABCDE";
        code39Field.AddStartStopChar = true;   // Add start/stop characters.
        code39Field.DisplayText = true;       // Show the encoded text below the barcode.

        // Update fields to render the new barcode.
        loadedDoc.UpdateFields();

        // Save the updated document.
        loadedDoc.Save("BarcodeDocument_WithSecondBarcode.docx");
    }
}
