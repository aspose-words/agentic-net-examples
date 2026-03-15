using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field. The second argument (true) updates the field immediately.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set minimal required properties for the barcode.
        // Here we use a simple QR code with a sample value.
        field.BarcodeType = "QR";
        field.BarcodeValue = "Sample123";

        // Save the document as DOCX.
        doc.Save("DisplayBarcode.docx");
    }
}
