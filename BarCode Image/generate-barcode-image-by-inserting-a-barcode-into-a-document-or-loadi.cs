using System;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Create a new document and insert a MERGEBARCODE field (QR code).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the field and configure its properties.
        FieldMergeBarcode mergeBarcode = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeBarcode.BarcodeType = "QR";                     // QR code type
        mergeBarcode.BarcodeValue = "ABC123";               // Data encoded in the QR code
        mergeBarcode.BackgroundColor = "F8BD69";            // Background colour (hex string without 0x)
        mergeBarcode.ForegroundColor = "B5413B";            // Foreground colour (hex string without 0x)
        mergeBarcode.ErrorCorrectionLevel = "3";           // QR error‑correction level (as string)
        mergeBarcode.ScalingFactor = "250";                // Scale 250 % (as string)
        mergeBarcode.SymbolHeight = "1000";                // Height in TWIPS (as string)
        mergeBarcode.SymbolRotation = "0";                 // No rotation (as string)

        builder.Writeln(); // Add a paragraph break after the barcode.

        // Update fields to render the barcode and save the document.
        doc.UpdateFields();
        doc.Save("BarcodeCreated.docx");

        // Load the previously saved document and add another barcode.
        Document loadedDoc = new Document("BarcodeCreated.docx");
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

        loadedBuilder.Writeln("Additional barcode:");

        // Insert a second MERGEBARCODE field with a different value.
        FieldMergeBarcode secondBarcode = (FieldMergeBarcode)loadedBuilder.InsertField(FieldType.FieldMergeBarcode, true);
        secondBarcode.BarcodeType = "QR";
        secondBarcode.BarcodeValue = "DEF456";
        secondBarcode.BackgroundColor = "F8BD69";
        secondBarcode.ForegroundColor = "B5413B";
        secondBarcode.ErrorCorrectionLevel = "3";
        secondBarcode.ScalingFactor = "250";
        secondBarcode.SymbolHeight = "1000";
        secondBarcode.SymbolRotation = "0";

        // Render the new field and save the updated document.
        loadedDoc.UpdateFields();
        loadedDoc.Save("BarcodeLoaded.docx");
    }
}
