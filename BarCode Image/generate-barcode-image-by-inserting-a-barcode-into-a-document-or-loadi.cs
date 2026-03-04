using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class BarcodeExample
{
    static void Main()
    {
        // -------------------- Create a document with a MERGEBARCODE field --------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field that will render a QR code.
        FieldMergeBarcode mergeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
        mergeField.BarcodeType = "QR";
        mergeField.BarcodeValue = "ABC123";
        mergeField.BackgroundColor = "0xF8BD69";
        mergeField.ForegroundColor = "0xB5413B";
        mergeField.ErrorCorrectionLevel = "3";
        mergeField.ScalingFactor = "250";
        mergeField.SymbolHeight = "1000";
        mergeField.SymbolRotation = "0";

        // Update the field so the barcode image is generated.
        mergeField.Update();

        // Save the document that contains the MERGEBARCODE field.
        doc.Save("BarcodeMerge.docx");

        // -------------------- Load the document and insert a barcode image --------------------
        Document loadedDoc = new Document("BarcodeMerge.docx");
        DocumentBuilder loadBuilder = new DocumentBuilder(loadedDoc);

        // Prepare barcode parameters identical to those used for the field.
        BarcodeParameters parameters = new BarcodeParameters
        {
            BarcodeType = "QR",
            BarcodeValue = "ABC123",
            BackgroundColor = "0xF8BD69",
            ForegroundColor = "0xB5413B",
            ErrorCorrectionLevel = "3",
            ScalingFactor = "250",
            SymbolHeight = "1000",
            SymbolRotation = "0"
        };

        // Generate the barcode image using the built‑in generator (or a custom one if assigned).
        using (Stream imgStream = loadedDoc.FieldOptions.BarcodeGenerator.GetBarcodeImage(parameters))
        {
            // Insert the generated image at the end of the document.
            loadBuilder.MoveToDocumentEnd();
            loadBuilder.InsertImage(imgStream);
        }

        // Save the final document that now contains the barcode image.
        loadedDoc.Save("BarcodeImage.docx");
    }
}
