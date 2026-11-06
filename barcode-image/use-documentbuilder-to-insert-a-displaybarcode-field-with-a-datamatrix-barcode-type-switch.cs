using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed insertion method.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "DATAMATRIX";          // Set the barcode type to DataMatrix.
        field.BarcodeValue = "1234567890";         // The data to encode.
        field.ForegroundColor = "0x000000";       // Optional: set foreground (black).
        field.BackgroundColor = "0xFFFFFF";       // Optional: set background (white).

        // Update fields so the barcode is rendered in the document.
        doc.UpdateFields();

        // Save the document to a DOCX file.
        doc.Save("DataMatrixBarcode.docx");
    }
}
