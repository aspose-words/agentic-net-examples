using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcode
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will generate a QR code.
        // The field is set to use the column "MyQRCode" from the mail‑merge data source.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        barcodeField.BarcodeType = "QR";          // QR code type.
        barcodeField.BarcodeValue = "MyQRCode";   // Name of the data‑source column.

        // Optional: customize QR code appearance (colors, error correction, scaling, etc.).
        barcodeField.BackgroundColor = "0xF8BD69";
        barcodeField.ForegroundColor = "0xB5413B";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        // Add a paragraph break after the field.
        builder.Writeln();

        // Prepare mail‑merge data: a single column that holds the values to encode.
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyQRCode");
        table.Rows.Add(new[] { "ABC123" });
        table.Rows.Add(new[] { "DEF456" });

        // Execute mail merge – each row creates a page with a DISPLAYBARCODE field
        // that renders the actual barcode image.
        doc.MailMerge.Execute(table);

        // Save the resulting document.
        doc.Save("BarcodeDocument.docx");
    }
}
