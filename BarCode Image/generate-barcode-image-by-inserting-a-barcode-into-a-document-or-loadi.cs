using System;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcode
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph break for clarity.
        builder.Writeln();

        // Insert a BARCODE field. The field will display a US ZIP code as a barcode.
        // The second argument (true) updates the field result immediately.
        FieldBarcode barcodeField = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);

        // Set the postal address (the data to encode). Here we use a simple ZIP code.
        barcodeField.PostalAddress = "12345";

        // Indicate that the address is a US postal address.
        barcodeField.IsUSPostalAddress = true;

        // Optionally, set a Facing Identification Mark (FIM) type.
        barcodeField.FacingIdentificationMark = "C";

        // Save the document in DOCX format.
        doc.Save("Barcode.docx");
    }
}
