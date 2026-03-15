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

        // Insert a DISPLAYBARCODE field that will render a QR code.
        // The field is inserted with custom colors, error correction level,
        // scaling factor, symbol height and rotation.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";
        barcodeField.BarcodeValue = "ABC123";
        barcodeField.BackgroundColor = "0xF8BD69";
        barcodeField.ForegroundColor = "0xB5413B";
        barcodeField.ErrorCorrectionLevel = "3";
        barcodeField.ScalingFactor = "250";
        barcodeField.SymbolHeight = "1000";
        barcodeField.SymbolRotation = "0";

        // Add a paragraph break after the field.
        builder.Writeln();

        // Save the document that now contains the barcode field.
        string insertedPath = Path.Combine(Environment.CurrentDirectory, "BarcodeInserted.docx");
        doc.Save(insertedPath);

        // Load the previously saved document.
        Document loadedDoc = new Document(insertedPath);
        DocumentBuilder loadedBuilder = new DocumentBuilder(loadedDoc);

        // Optionally update fields to ensure the barcode image is generated.
        loadedDoc.UpdateFields();

        // Save the loaded document under a different name.
        string loadedPath = Path.Combine(Environment.CurrentDirectory, "BarcodeLoaded.docx");
        loadedDoc.Save(loadedPath);
    }
}
