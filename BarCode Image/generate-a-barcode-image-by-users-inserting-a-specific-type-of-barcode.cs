using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a BARCODE field. The second argument (true) updates the field result immediately.
        FieldBarcode barcodeField = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);

        // Set the data to encode – a valid U.S. ZIP code.
        barcodeField.PostalAddress = "96801";

        // Indicate that the value is a U.S. postal address so the field generates a postal barcode.
        barcodeField.IsUSPostalAddress = true;

        // Optional: add a Facing Identification Mark (FIM) to the barcode.
        barcodeField.FacingIdentificationMark = "C";

        // Update the field to ensure the barcode image is generated.
        barcodeField.Update();

        // Save the document as a DOCX file.
        doc.Save("BarcodeDocument.docx");
    }
}
