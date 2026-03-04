using System;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcode
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field. This field will render a barcode directly.
        // The second argument (true) tells the builder to update the field result immediately.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);

        // Configure the barcode type and the value to encode.
        barcodeField.BarcodeType = "CODE39";          // Choose barcode type (e.g., CODE39, QR, EAN13, etc.).
        barcodeField.BarcodeValue = "12345ABCDE";    // The data that will be encoded.
        barcodeField.AddStartStopChar = true;        // For CODE39, include start/stop characters.

        // Update all fields in the document to ensure the barcode is rendered.
        doc.UpdateFields();

        // Save the document in DOCX format.
        doc.Save("Barcode.docx");
    }
}
