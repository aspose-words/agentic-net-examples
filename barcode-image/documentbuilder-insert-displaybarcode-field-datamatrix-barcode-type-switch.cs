using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field at the current cursor position.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Configure the field to display a DataMatrix barcode.
        field.BarcodeType = "DATAMATRIX";   // Barcode type switch.
        field.BarcodeValue = "ABC123";      // The data to encode.
        field.ScalingFactor = "250";        // Optional: adjust size.

        // Add a line break after the field for readability.
        builder.Writeln();

        // Save the document to disk.
        doc.Save("DisplayBarcodeDataMatrix.docx");
    }
}
