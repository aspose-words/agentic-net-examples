using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new document and insert a QR barcode.
        // -------------------------------------------------
        Document doc = new Document();                     // create a blank document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field and configure it as a QR code.
        FieldDisplayBarcode qrField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        qrField.BarcodeType = "QR";
        qrField.BarcodeValue = "ABC123";
        qrField.BackgroundColor = "0xF8BD69";
        qrField.ForegroundColor = "0xB5413B";
        qrField.ErrorCorrectionLevel = "3";
        qrField.ScalingFactor = "250";
        qrField.SymbolHeight = "1000";
        qrField.SymbolRotation = "0";

        builder.Writeln(); // add a paragraph break after the field

        // Save the newly created document as DOCX.
        doc.Save("BarcodeDocument.docx");

        // -------------------------------------------------
        // 2. Load an existing DOCX and add a CODE39 barcode.
        // -------------------------------------------------
        Document existingDoc = new Document("Existing.docx"); // load an existing document
        DocumentBuilder existingBuilder = new DocumentBuilder(existingDoc);

        // Insert a DISPLAYBARCODE field and configure it as a CODE39 barcode.
        FieldDisplayBarcode code39Field = (FieldDisplayBarcode)existingBuilder.InsertField(FieldType.FieldDisplayBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "12345ABCDE";
        code39Field.AddStartStopChar = true;

        existingBuilder.Writeln(); // optional line break

        // Save the modified document.
        existingDoc.Save("ExistingWithBarcode.docx");
    }
}
