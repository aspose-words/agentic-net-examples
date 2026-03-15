using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a new document and insert a DISPLAYBARCODE field.
        // -----------------------------------------------------------------
        Document doc = new Document();                     // create a blank document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field that will render a QR code.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "QR";                   // type of barcode
        barcodeField.BarcodeValue = "ABC123";              // data to encode
        barcodeField.BackgroundColor = "0xF8BD69";         // optional background colour
        barcodeField.ForegroundColor = "0xB5413B";         // optional foreground colour
        barcodeField.ErrorCorrectionLevel = "3";           // QR error correction level
        barcodeField.ScalingFactor = "250";                // scaling factor (percentage)
        barcodeField.SymbolHeight = "1000";                // height in TWIPS
        barcodeField.SymbolRotation = "0";                 // rotation

        // Add a paragraph break after the field for readability.
        builder.Writeln();

        // Save the newly created document as DOCX.
        doc.Save("BarcodeInserted.docx");

        // -----------------------------------------------------------------
        // 2. Load an existing document and add a MERGEBARCODE field.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document("BarcodeInserted.docx"); // load the previously saved file
        DocumentBuilder loadBuilder = new DocumentBuilder(loadedDoc);

        // Move the cursor to the end of the document.
        loadBuilder.MoveToDocumentEnd();

        // Insert a MERGEBARCODE field that can be used in a mail‑merge scenario.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)loadBuilder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "CODE39";
        mergeField.BarcodeValue = "MyCODE39Column"; // name of the data source column
        mergeField.AddStartStopChar = true;         // include start/stop characters

        // Save the modified document.
        loadedDoc.Save("BarcodeMerged.docx");
    }
}
