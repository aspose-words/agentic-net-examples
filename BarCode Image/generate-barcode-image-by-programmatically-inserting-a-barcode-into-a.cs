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

        // -------------------------------------------------
        // Insert a MERGEBARCODE field.
        // This field will be used during a mail merge to generate barcodes.
        // -------------------------------------------------
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "QR";                 // Type of barcode.
        mergeField.BarcodeValue = "MyQRCode";          // Name of the data source column.
        mergeField.BackgroundColor = "0xF8BD69";       // Custom background colour.
        mergeField.ForegroundColor = "0xB5413B";       // Custom foreground colour.
        mergeField.ErrorCorrectionLevel = "3";         // QR error correction level.
        mergeField.ScalingFactor = "250";              // Scale the symbol.
        mergeField.SymbolHeight = "1000";              // Height in TWIPS.
        mergeField.SymbolRotation = "0";               // No rotation.

        // Add a line break after the MERGEBARCODE field.
        builder.Writeln();

        // -------------------------------------------------
        // Insert a DISPLAYBARCODE field.
        // This field directly displays a barcode without mail merge.
        // -------------------------------------------------
        FieldDisplayBarcode displayField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        displayField.BarcodeType = "CODE39";
        displayField.BarcodeValue = "12345ABCDE";
        displayField.AddStartStopChar = true;          // Show start/stop characters.

        // Add another line break.
        builder.Writeln();

        // -------------------------------------------------
        // Prepare a simple data source for the MERGEBARCODE field.
        // The column name must match the BarcodeValue set above ("MyQRCode").
        // -------------------------------------------------
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyQRCode");
        table.Rows.Add(new object[] { "ABC123" });
        table.Rows.Add(new object[] { "DEF456" });

        // Execute mail merge – each row creates a new page with a DISPLAYBARCODE field.
        doc.MailMerge.Execute(table);

        // -------------------------------------------------
        // Save the document in DOCX format.
        // -------------------------------------------------
        string outputPath = "MergeBarcodeDisplayBarcode.docx";
        doc.Save(outputPath);
    }
}
