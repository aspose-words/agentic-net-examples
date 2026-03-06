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

        // Insert a MERGEBARCODE field. This field will be replaced by a DISPLAYBARCODE
        // field during mail merge, allowing us to generate a barcode for each data row.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "QR";                     // Type of barcode.
        mergeField.BarcodeValue = "MyQRCode";              // Name of the column that holds the value.

        // Optional visual customizations for the QR code.
        // The color properties expect a string with a hex RGB value (without the leading "0x").
        mergeField.BackgroundColor = "F8BD69";            // Background colour (hex RGB).
        mergeField.ForegroundColor = "B5413B";            // Foreground colour (hex RGB).
        mergeField.ErrorCorrectionLevel = "3";            // QR error correction level (0‑3).
        mergeField.ScalingFactor = "250";                 // Scale as a percentage.
        mergeField.SymbolHeight = "1000";                 // Height in TWIPS (1/1440 inch).
        mergeField.SymbolRotation = "0";                  // No rotation.

        // Prepare a data source for mail merge.
        DataTable data = new DataTable("Barcodes");
        data.Columns.Add("MyQRCode");
        data.Rows.Add("ABC123");
        data.Rows.Add("DEF456");

        // Perform mail merge – each row creates a new page with a barcode.
        doc.MailMerge.Execute(data);

        // Save the resulting document as DOCX.
        doc.Save("BarcodeDocument.docx");
    }
}
