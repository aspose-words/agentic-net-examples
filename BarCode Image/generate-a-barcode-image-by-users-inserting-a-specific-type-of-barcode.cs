using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field. This field type allows setting barcode parameters directly.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, false);

        // Configure the barcode type and value.
        barcodeField.BarcodeType = "QR";          // QR code
        barcodeField.BarcodeValue = "ABC123";    // Data to encode

        // Optionally, display the encoded text below the barcode.
        barcodeField.DisplayText = true;

        // Save the document as DOCX.
        doc.Save("BarcodeDocument.docx");
    }
}
