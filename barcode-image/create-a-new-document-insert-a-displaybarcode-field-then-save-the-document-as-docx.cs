using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed API.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";          // Set the barcode type.
        barcodeField.BarcodeValue = "1234567890"; // Set the data to encode.

        // Update fields so the field result is generated.
        doc.UpdateFields();

        // Save the document as DOCX.
        doc.Save("DisplayBarcode.docx");
    }
}
