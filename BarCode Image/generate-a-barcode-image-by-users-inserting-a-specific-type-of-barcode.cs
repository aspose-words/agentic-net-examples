using System;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcodeWithBarField
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a BARCODE field. This field is intended for U.S. postal barcodes.
        // The field will display the ZIP code 96801 as a barcode.
        FieldBarcode barcodeField = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);

        // Set the ZIP code value.
        barcodeField.PostalAddress = "96801";

        // Indicate that the value is a U.S. postal address.
        barcodeField.IsUSPostalAddress = true;

        // Optional: add a Facing Identification Mark (FIM) to the barcode.
        barcodeField.FacingIdentificationMark = "C";

        // Save the document as a DOCX file.
        doc.Save("BarcodeBarField.docx");
    }
}
