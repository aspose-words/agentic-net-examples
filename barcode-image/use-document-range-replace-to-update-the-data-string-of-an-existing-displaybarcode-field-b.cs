using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with a placeholder value.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "PLACEHOLDER";
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";

        // Replace the placeholder value using Document.Range.Replace.
        FindReplaceOptions options = new FindReplaceOptions();
        doc.Range.Replace("PLACEHOLDER", "NEWVALUE", options);

        // Update fields to reflect the new barcode value.
        doc.UpdateFields();

        // Save the document.
        doc.Save("output.docx");
    }
}
