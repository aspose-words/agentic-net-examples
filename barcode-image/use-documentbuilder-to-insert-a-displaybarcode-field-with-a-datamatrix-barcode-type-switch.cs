using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field using the typed insertion method.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // Set the barcode type to DataMatrix and provide the barcode value.
        field.BarcodeType = "DATAMATRIX";
        field.BarcodeValue = "1234567890";

        // Optional: set foreground and background colors.
        field.ForegroundColor = "0x000000"; // Black
        field.BackgroundColor = "0xFFFFFF"; // White

        // Update fields to ensure the barcode is rendered in the document.
        doc.UpdateFields();

        // Save the document to a DOCX file.
        doc.Save("DataMatrixBarcode.docx");
    }
}
