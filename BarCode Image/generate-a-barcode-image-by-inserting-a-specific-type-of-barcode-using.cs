using System;
using Aspose.Words;
using Aspose.Words.Fields;

class GenerateBarcode
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Obtain the first paragraph (or create one) to host the barcode field.
        Paragraph para = doc.FirstSection.Body.FirstParagraph ?? doc.FirstSection.Body.AppendParagraph(string.Empty);

        // Build a MERGEBARCODE field that will display a CODE39 barcode with the value "12345ABCDE".
        // The field code syntax for MERGEBARCODE uses switches:
        //   \b – barcode type
        //   \d – barcode data (value)
        //   \a – add start/stop characters (optional, true/false)
        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldMergeBarcode);
        fieldBuilder.AddSwitch("\\b", "CODE39");          // Set barcode type.
        fieldBuilder.AddSwitch("\\d", "12345ABCDE");     // Set barcode value.
        fieldBuilder.AddSwitch("\\a", "true");           // Add start/stop characters (string, not bool).

        // Insert the field at the end of the paragraph.
        Field barcodeField = fieldBuilder.BuildAndInsert(para);

        // Update the field so that the barcode image is generated.
        doc.UpdateFields();

        // Save the document containing the barcode.
        doc.Save("BarcodeDocument.docx");
    }
}
