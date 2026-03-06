using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the primary header so the barcode will appear on every page
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert a DISPLAYBARCODE field. The second argument (true) tells the builder to insert the field code.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);

        // All properties of FieldDisplayBarcode are strings, so assign string values.
        // Hex colour values must be supplied without the "0x" prefix.
        barcodeField.BarcodeType = "QR";               // Type of barcode
        barcodeField.BarcodeValue = "ABC123";          // Data encoded in the barcode
        barcodeField.BackgroundColor = "F8BD69";       // Background colour (hex)
        barcodeField.ForegroundColor = "B5413B";       // Foreground colour (hex)
        barcodeField.ErrorCorrectionLevel = "3";       // QR error‑correction level (0‑3)
        barcodeField.ScalingFactor = "250";            // Scale factor as string
        barcodeField.SymbolHeight = "1000";            // Height in twips as string
        barcodeField.SymbolRotation = "0";             // Rotation in degrees as string

        // Optionally add a line break after the barcode in the header
        builder.Writeln();

        // Save the document to disk
        doc.Save("BarcodeOnEachPage.docx");
    }
}
