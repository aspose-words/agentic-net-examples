using System;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // Folder where the generated documents will be saved.
        string artifactsDir = @"C:\Artifacts\";

        // -------------------------------------------------
        // 1. Create a new document and insert barcode fields.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a QR code with custom colors and scaling.
        FieldDisplayBarcode qrField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        qrField.BarcodeType = "QR";
        qrField.BarcodeValue = "ABC123";
        qrField.BackgroundColor = "0xF8BD69";
        qrField.ForegroundColor = "0xB5413B";
        qrField.ErrorCorrectionLevel = "3";
        qrField.ScalingFactor = "250";
        qrField.SymbolHeight = "1000";
        qrField.SymbolRotation = "0";

        builder.Writeln(); // separate the barcodes

        // Insert an EAN13 barcode that displays the numeric value below the bars.
        FieldDisplayBarcode eanField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        eanField.BarcodeType = "EAN13";
        eanField.BarcodeValue = "501234567890";
        eanField.DisplayText = true;
        eanField.PosCodeStyle = "CASE";
        eanField.FixCheckDigit = true;

        // Save the newly created document.
        doc.Save(artifactsDir + "BarcodeDocument.docx");

        // -------------------------------------------------
        // 2. Load the saved document and add another barcode.
        // -------------------------------------------------
        Document loadedDoc = new Document(artifactsDir + "BarcodeDocument.docx");
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

        // Insert a CODE39 barcode with start/stop characters.
        FieldDisplayBarcode code39Field = (FieldDisplayBarcode)loadedBuilder.InsertField(FieldType.FieldDisplayBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "12345ABCDE";
        code39Field.AddStartStopChar = true;

        // Save the document after adding the new barcode.
        loadedDoc.Save(artifactsDir + "BarcodeDocument_WithCode39.docx");
    }
}
