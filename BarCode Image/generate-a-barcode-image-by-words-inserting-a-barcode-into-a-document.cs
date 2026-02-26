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

        // Insert a MERGEBARCODE field that will generate a QR code.
        FieldMergeBarcode mergeBarcode = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeBarcode.BarcodeType = "QR";
        mergeBarcode.BarcodeValue = "Sample123";

        // Optional appearance settings – colors must be supplied as strings (hex values without the 0x prefix).
        mergeBarcode.BackgroundColor = "F8BD69"; // Light orange background.
        mergeBarcode.ForegroundColor = "B5413B"; // Dark red bars.
        mergeBarcode.ErrorCorrectionLevel = "3";
        mergeBarcode.ScalingFactor = "250";
        mergeBarcode.SymbolHeight = "1000";
        mergeBarcode.SymbolRotation = "0";

        // Add a paragraph break after the field.
        builder.Writeln();

        // Save the document as DOCX.
        doc.Save("BarcodeMerge.docx");

        // -----------------------------------------------------------------
        // Load the saved document and add a DISPLAYBARCODE field that shows a CODE39 barcode.
        Document loadedDoc = new Document("BarcodeMerge.docx");
        DocumentBuilder loadBuilder = new DocumentBuilder(loadedDoc);
        loadBuilder.MoveToDocumentEnd();

        // Insert a DISPLAYBARCODE field.
        FieldDisplayBarcode displayBarcode = (FieldDisplayBarcode)loadBuilder.InsertField(FieldType.FieldDisplayBarcode, true);
        displayBarcode.BarcodeType = "CODE39";
        displayBarcode.BarcodeValue = "12345ABCDE";
        displayBarcode.AddStartStopChar = true; // Show start/stop characters.

        // Save the modified document.
        loadedDoc.Save("BarcodeDisplay.docx");
    }
}
