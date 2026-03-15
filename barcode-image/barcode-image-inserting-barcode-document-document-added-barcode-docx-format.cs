using Aspose.Words;
using Aspose.Words.Fields;
using System.Data;

class BarcodeExample
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new document and insert a DISPLAYBARCODE field.
        // -------------------------------------------------
        Document doc = new Document();                     // create a blank document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field that will render a QR code.
        FieldDisplayBarcode displayField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        displayField.BarcodeType = "QR";                   // QR code type
        displayField.BarcodeValue = "ABC123";              // data to encode
        displayField.BackgroundColor = "0xF8BD69";         // custom background color
        displayField.ForegroundColor = "0xB5413B";         // custom foreground color
        displayField.ErrorCorrectionLevel = "3";           // highest error correction
        displayField.ScalingFactor = "250";                // 250 %
        displayField.SymbolHeight = "1000";                // height in TWIPS
        displayField.SymbolRotation = "0";                 // no rotation

        builder.Writeln(); // add a paragraph break after the field

        // Save the document containing the barcode field.
        doc.Save("BarcodeDisplay.docx");

        // -------------------------------------------------
        // 2. Create a document with a MERGEBARCODE field and perform mail merge.
        // -------------------------------------------------
        Document mergeDoc = new Document();                // create another blank document
        DocumentBuilder mergeBuilder = new DocumentBuilder(mergeDoc);

        // Insert a MERGEBARCODE field that will be populated during mail merge.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)mergeBuilder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "CODE39";                 // CODE39 barcode
        mergeField.BarcodeValue = "MyCodeColumn";          // column name in data source
        mergeField.AddStartStopChar = true;                // include start/stop characters

        // Prepare a DataTable that mimics a data source.
        DataTable table = new DataTable("Barcodes");
        table.Columns.Add("MyCodeColumn");
        table.Rows.Add(new object[] { "12345ABCDE" });
        table.Rows.Add(new object[] { "67890FGHIJ" });

        // Execute mail merge – each row creates a DISPLAYBARCODE field.
        mergeDoc.MailMerge.Execute(table);

        // Save the merged document with generated barcodes.
        mergeDoc.Save("BarcodeMerge.docx");
    }
}
