using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a new DOCX and insert a MERGEBARCODE field (QR code)
        // -------------------------------------------------
        Document doc = new Document();                     // create a blank document
        DocumentBuilder builder = new DocumentBuilder(doc); // helper to add content

        // Insert a MERGEBARCODE field that will generate a QR code.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "QR";                     // type of barcode
        mergeField.BarcodeValue = "ABC123";                // data to encode

        // Optional visual customizations
        mergeField.BackgroundColor = "0xF8BD69";           // background colour
        mergeField.ForegroundColor = "0xB5413B";           // foreground colour
        mergeField.ErrorCorrectionLevel = "3";             // QR error correction
        mergeField.ScalingFactor = "250";                  // scaling factor (percent)
        mergeField.SymbolHeight = "1000";                  // height in TWIPS
        mergeField.SymbolRotation = "0";                   // rotation

        builder.Writeln(); // add a paragraph break after the field

        // Save the document containing the MERGEBARCODE field.
        doc.Save("BarcodeMerge.docx");

        // -------------------------------------------------
        // 2. Load the saved document and insert a BARCODE field (US ZIP code)
        // -------------------------------------------------
        Document loadedDoc = new Document("BarcodeMerge.docx"); // load existing DOCX
        DocumentBuilder loadBuilder = new DocumentBuilder(loadedDoc);
        loadBuilder.MoveToDocumentEnd(); // position cursor at the end

        // Insert a BARCODE field that displays a US postal barcode.
        FieldBarcode barcodeField = (FieldBarcode)loadBuilder.InsertField(FieldType.FieldBarcode, true);
        barcodeField.PostalAddress = "12345"; // ZIP code to encode
        barcodeField.IsUSPostalAddress = true; // indicate US postal address
        barcodeField.FacingIdentificationMark = "C"; // optional FIM marker

        // Save the updated document.
        loadedDoc.Save("BarcodeLoaded.docx");
    }
}
