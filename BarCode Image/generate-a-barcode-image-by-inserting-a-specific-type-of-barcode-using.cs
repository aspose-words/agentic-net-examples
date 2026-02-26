using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a BARCODE field that will render a barcode image.
        FieldBarcode barcodeField = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);

        // Set the value to be encoded (e.g., a US ZIP code).
        barcodeField.PostalAddress = "12345";

        // Specify that the value is a US postal address so the field knows which barcode symbology to use.
        barcodeField.IsUSPostalAddress = true;

        // Optional: add a Facing Identification Mark (FIM) to the barcode.
        barcodeField.FacingIdentificationMark = "C";

        // Force the field to calculate its result and generate the barcode image.
        barcodeField.Update();

        // Save the document containing the barcode.
        doc.Save("BarcodeDocument.docx");
    }
}
