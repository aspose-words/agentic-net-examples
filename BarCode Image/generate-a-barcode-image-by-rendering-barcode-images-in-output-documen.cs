using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class BarcodeToPdfExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // Insert a MERGEBARCODE field for a QR code.
        // -------------------------------------------------
        FieldMergeBarcode qrField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        qrField.BarcodeType = "QR";
        qrField.BarcodeValue = "ABC123";
        qrField.BackgroundColor = "0xF8BD69";
        qrField.ForegroundColor = "0xB5413B";
        qrField.ErrorCorrectionLevel = "3";
        qrField.ScalingFactor = "250";
        qrField.SymbolHeight = "1000";
        qrField.SymbolRotation = "0";

        // Add a line break between barcodes.
        builder.Writeln();

        // -------------------------------------------------
        // Insert a MERGEBARCODE field for a CODE39 barcode.
        // -------------------------------------------------
        FieldMergeBarcode code39Field = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        code39Field.BarcodeType = "CODE39";
        code39Field.BarcodeValue = "12345ABCDE";
        code39Field.AddStartStopChar = true;

        // -------------------------------------------------
        // Update fields so that barcode images are generated.
        // -------------------------------------------------
        doc.UpdateFields();

        // -------------------------------------------------
        // Save the document as PDF.
        // -------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save("BarcodesOutput.pdf", pdfOptions);
    }
}
