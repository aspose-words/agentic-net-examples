using System;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field that will render a QR code.
        FieldDisplayBarcode qrField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        qrField.BarcodeType = "QR";
        qrField.BarcodeValue = "ABC123";
        qrField.BackgroundColor = "0xF8BD69";
        qrField.ForegroundColor = "0xB5413B";
        qrField.ErrorCorrectionLevel = "3";
        qrField.ScalingFactor = "250";
        qrField.SymbolHeight = "1000";
        qrField.SymbolRotation = "0";

        // Add a paragraph break after the field.
        builder.Writeln();

        // Save the document that now contains the QR barcode.
        doc.Save("Barcode_QR.docx");

        // Load the previously saved document.
        Document loadedDoc = new Document("Barcode_QR.docx");
        DocumentBuilder loadBuilder = new DocumentBuilder(loadedDoc);

        // Insert another DISPLAYBARCODE field, this time a CODE39 barcode with start/stop characters.
        FieldDisplayBarcode code39Field = (FieldDisplayBarcode)loadBuilder.InsertField(FieldType.FieldDisplayBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "12345ABCDE";
        code39Field.AddStartStopChar = true;

        // Save the updated document containing both barcodes.
        loadedDoc.Save("Barcode_QR_and_CODE39.docx");
    }
}
