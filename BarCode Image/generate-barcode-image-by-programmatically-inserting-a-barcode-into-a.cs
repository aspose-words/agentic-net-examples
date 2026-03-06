using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class MergeBarcodeExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field. This field will be replaced by a DISPLAYBARCODE
        // field for each row of the mail‑merge data source.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "QR";               // Type of barcode to generate.
        mergeField.BarcodeValue = "MyQRCode";        // Name of the column in the data source.

        // Optional: customize appearance of the QR code.
        mergeField.BackgroundColor = "0xF8BD69";
        mergeField.ForegroundColor = "0xB5413B";
        mergeField.ErrorCorrectionLevel = "3";
        mergeField.ScalingFactor = "250";
        mergeField.SymbolHeight = "1000";
        mergeField.SymbolRotation = "0";

        // Add a line break after the field so each barcode appears on its own line.
        builder.Writeln();

        // Prepare a DataTable that will serve as the mail‑merge data source.
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyQRCode");
        table.Rows.Add(new[] { "ABC123" });
        table.Rows.Add(new[] { "DEF456" });

        // Execute the mail merge. Each row creates a new page with a DISPLAYBARCODE field.
        doc.MailMerge.Execute(table);

        // Save the resulting document in DOCX format.
        doc.Save("MergeBarcode.docx");
    }
}
