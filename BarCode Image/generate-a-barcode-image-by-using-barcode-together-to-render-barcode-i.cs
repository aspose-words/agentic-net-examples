using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Saving;

class BarcodeToPdf
{
    static void Main()
    {
        // Paths for the output files.
        string docxPath = "BarcodeDocument.docx";
        string pdfPath = "BarcodeDocument.pdf";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field. The second argument (true) tells the builder to
        // add the field to the end of the current paragraph.
        builder.InsertField(FieldType.FieldMergeBarcode, true);

        // Retrieve the field we just inserted (it will be the last field in the document).
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)doc.Range.Fields[doc.Range.Fields.Count - 1];

        // Configure the barcode parameters.
        barcodeField.BarcodeType = "QR";          // QR code type.
        barcodeField.BarcodeValue = "ABC123";    // Data to encode.
        barcodeField.DisplayText = true;         // Show the encoded text below the image.

        // Update fields so the barcode image is generated.
        doc.UpdateFields();

        // Save the document as DOCX (optional, shows the field in Word format).
        doc.Save(docxPath);

        // Save the same document as PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        doc.Save(pdfPath, pdfOptions);
    }
}
