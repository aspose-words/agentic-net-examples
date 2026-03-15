using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field for a QR code.
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

        // Save the document containing the QR barcode.
        string initialPath = Path.Combine(Environment.CurrentDirectory, "BarcodeDocument.docx");
        doc.Save(initialPath);

        // Load the saved document and append another barcode (CODE39) at the end.
        Document loadedDoc = new Document(initialPath);
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);
        loadedBuilder.MoveToDocumentEnd();

        FieldDisplayBarcode code39Field = (FieldDisplayBarcode)loadedBuilder.InsertField(FieldType.FieldDisplayBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "12345ABCDE";
        code39Field.AddStartStopChar = true;

        // Save the updated document.
        string updatedPath = Path.Combine(Environment.CurrentDirectory, "BarcodeDocument_Updated.docx");
        loadedDoc.Save(updatedPath);
    }
}
