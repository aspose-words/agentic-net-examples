using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a MERGEBARCODE field that will generate a QR code with the value "ABC123".
        // The field code will look like: MERGEBARCODE QR "ABC123"
        FieldBuilder barcodeField = new FieldBuilder(FieldType.FieldMergeBarcode);
        barcodeField.AddArgument("QR");        // Barcode type.
        barcodeField.AddArgument("ABC123");    // Barcode value.

        // Insert the field at the current cursor position (end of the current paragraph).
        barcodeField.BuildAndInsert(builder.CurrentParagraph);

        // Update fields so that the barcode image is generated.
        // If a custom IBarcodeGenerator is not supplied, the field will display a placeholder.
        doc.UpdateFields();

        // Save the document in DOCX format.
        doc.Save("BarcodeDocument.docx");
    }
}
