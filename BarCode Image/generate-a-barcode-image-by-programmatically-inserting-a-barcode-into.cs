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

        // Insert a MERGEBARCODE field. The field will be populated from a data source during mail merge.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "QR";               // QR code type
        mergeField.BarcodeValue = "MyQRCode";        // Column name in the data source
        // Optional visual customizations.
        mergeField.BackgroundColor = "0xF8BD69";
        mergeField.ForegroundColor = "0xB5413B";
        mergeField.ErrorCorrectionLevel = "3";
        mergeField.ScalingFactor = "250";
        mergeField.SymbolHeight = "1000";
        mergeField.SymbolRotation = "0";

        // Add a paragraph break after the MERGEBARCODE field.
        builder.Writeln();

        // Prepare a DataTable that matches the MERGEBARCODE field's BarcodeValue column.
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyQRCode");
        table.Rows.Add(new object[] { "ABC123" });
        table.Rows.Add(new object[] { "DEF456" });

        // Perform mail merge. Each row creates a DISPLAYBARCODE field that displays the QR code.
        doc.MailMerge.Execute(table);

        // Insert a DISPLAYBARCODE field directly (without mail merge) as an additional example.
        builder.Writeln(); // separate paragraph
        FieldDisplayBarcode displayField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        displayField.BarcodeType = "CODE39";
        displayField.BarcodeValue = "12345ABCDE";
        displayField.AddStartStopChar = true; // show start/stop characters

        // Save the resulting DOCX document.
        doc.Save("BarcodeDocument.docx");
    }
}
