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

        // Insert a MERGEBARCODE field. The field will be replaced by a DISPLAYBARCODE
        // field for each row of the data source during mail merge.
        // The field is set to read the value from the column "MyQRCode".
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "QR";          // QR code type
        mergeField.BarcodeValue = "MyQRCode";   // column name in the data source

        // Optional: customize appearance (background, foreground, error correction, etc.).
        mergeField.BackgroundColor = "0xF8BD69";
        mergeField.ForegroundColor = "0xB5413B";
        mergeField.ErrorCorrectionLevel = "3";
        mergeField.ScalingFactor = "250";
        mergeField.SymbolHeight = "1000";
        mergeField.SymbolRotation = "0";

        // Add a line break after the field so each barcode appears on its own line.
        builder.Writeln();

        // Prepare a data table that will be used for mail merge.
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyQRCode");
        table.Rows.Add(new object[] { "ABC123" });
        table.Rows.Add(new object[] { "DEF456" });

        // Perform mail merge – each row creates a DISPLAYBARCODE field with the QR code.
        doc.MailMerge.Execute(table);

        // Save the resulting document as DOCX.
        doc.Save("BarcodeMerge.docx");
    }
}
