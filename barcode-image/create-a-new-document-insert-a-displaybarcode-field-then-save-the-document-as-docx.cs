using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "1234567890";
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";

        // Update fields to render the barcode.
        doc.UpdateFields();

        // Save the document as DOCX.
        doc.Save("output.docx");
    }
}
