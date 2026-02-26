using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will be filled by mail merge.
        // The field will generate a QR code.
        FieldMergeBarcode mergeBarcode = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeBarcode.BarcodeType = "QR";
        mergeBarcode.BarcodeValue = "MyQRCode";

        // Optional visual customizations – all properties are strings.
        mergeBarcode.BackgroundColor = "F8BD69";      // Background color (hex RGB, without 0x).
        mergeBarcode.ForegroundColor = "B5413B";      // Foreground (bars) color.
        mergeBarcode.ErrorCorrectionLevel = "3";      // QR error correction level (0‑3).
        mergeBarcode.ScalingFactor = "250";           // Scale to 250 %.
        mergeBarcode.SymbolHeight = "1000";           // Height in TWIPS.
        mergeBarcode.SymbolRotation = "0";            // No rotation.

        builder.Writeln(); // Add a paragraph break after the field.

        // Prepare a data source for mail merge.
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyQRCode");
        table.Rows.Add("ABC123");
        table.Rows.Add("DEF456");

        // Perform mail merge – each row creates a new page with a DISPLAYBARCODE field.
        doc.MailMerge.Execute(table);

        // Save the document that now contains barcode fields.
        doc.Save("BarcodeMerge.docx");

        // -----------------------------------------------------------------
        // Load the previously saved document and update fields to render the barcodes.
        Document loadedDoc = new Document("BarcodeMerge.docx");
        loadedDoc.UpdateFields(); // Forces field calculation and image generation.
        loadedDoc.Save("BarcodeMergedUpdated.docx");
    }
}
