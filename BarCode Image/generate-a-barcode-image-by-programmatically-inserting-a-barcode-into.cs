using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field. This field will be replaced by a DISPLAYBARCODE
        // field for each row of the mail‑merge data source.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "QR";               // Type of barcode.
        mergeField.BarcodeValue = "MyQRCode";        // Column name in the data source.
        // Optional visual customisation.
        mergeField.BackgroundColor = "0xF8BD69";
        mergeField.ForegroundColor = "0xB5413B";
        mergeField.ErrorCorrectionLevel = "3";
        mergeField.ScalingFactor = "250";
        mergeField.SymbolHeight = "1000";
        mergeField.SymbolRotation = "0";

        // Add a paragraph break after the field.
        builder.Writeln();

        // Prepare a simple data table with a column that matches the MERGEBARCODE field.
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyQRCode");
        table.Rows.Add(new[] { "ABC123" });
        table.Rows.Add(new[] { "DEF456" });

        // Execute mail merge – each row creates a new page containing a DISPLAYBARCODE field.
        doc.MailMerge.Execute(table);

        // Verify that the fields have been converted to DISPLAYBARCODE.
        // (Optional – can be removed in production code.)
        foreach (Field field in doc.Range.Fields)
        {
            Console.WriteLine($"{field.Type}: {field.GetFieldCode()}");
        }

        // Save the resulting document as DOCX.
        doc.Save("BarcodeDocument.docx");
    }
}
