using System;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeToPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph heading.
        builder.Writeln("Sample barcode (U.S. ZIP code):");

        // Insert a BARCODE field that will render the ZIP code as a barcode.
        // The field will display the barcode for the value set in PostalAddress.
        FieldBarcode barcodeField = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);
        barcodeField.PostalAddress = "96801";          // ZIP code to encode.
        barcodeField.IsUSPostalAddress = true;        // Specify that it is a U.S. postal address.
        barcodeField.FacingIdentificationMark = "C";  // Optional FIM character.

        // Update fields so the barcode image is generated.
        doc.UpdateFields();

        // Save the document as PDF.
        doc.Save("BarcodeOutput.pdf", SaveFormat.Pdf);
    }
}
