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

        // Insert a MERGEBARCODE field that will generate a QR code.
        // The second argument (true) tells the builder to update the field immediately.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        barcodeField.BarcodeType = "QR";          // Set the barcode type.
        barcodeField.BarcodeValue = "HelloWorld"; // Set the data to encode.

        // Ensure all fields are up‑to‑date (necessary if you added fields without auto‑updating).
        doc.UpdateFields();

        // Save the document containing the generated barcode.
        doc.Save("BarcodeDocument.docx");
    }
}
