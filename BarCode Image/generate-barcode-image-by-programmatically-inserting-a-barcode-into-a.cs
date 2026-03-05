using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field. This field will be populated during mail merge.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "QR";               // Type of barcode.
        mergeField.BarcodeValue = "MyQRCode";        // Name of the column in the data source.
        // Optional visual customizations.
        mergeField.BackgroundColor = "0xF8BD69";
        mergeField.ForegroundColor = "0xB5413B";
        mergeField.ErrorCorrectionLevel = "3";
        mergeField.ScalingFactor = "250";
        mergeField.SymbolHeight = "1000";
        mergeField.SymbolRotation = "0";

        builder.Writeln(); // Move to the next line.

        // Insert a DISPLAYBARCODE field directly (no mail merge needed).
        FieldDisplayBarcode displayField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        displayField.BarcodeType = "CODE39";
        displayField.BarcodeValue = "12345ABCDE";
        displayField.AddStartStopChar = true; // Show start/stop characters for CODE39.

        builder.Writeln(); // Add a line break after the field.

        // Prepare a data source for mail merge.
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyQRCode");
        table.Rows.Add("ABC123");
        table.Rows.Add("DEF456");

        // Perform mail merge. Each row creates a new page with a DISPLAYBARCODE field.
        doc.MailMerge.Execute(table);

        // Save the document in DOCX format.
        doc.Save("BarcodeDemo.docx");
    }
}
