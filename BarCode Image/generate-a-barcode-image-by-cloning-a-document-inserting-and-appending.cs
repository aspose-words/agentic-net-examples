using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsBarcodeDemo
{
    class Program
    {
        static void Main()
        {
            // Load the original DOCX document.
            Document originalDoc = new Document("Source.docx");

            // Clone the original document – this creates a deep copy.
            Document clonedDoc = originalDoc.Clone();

            // Insert a MERGEBARCODE field into the cloned document.
            DocumentBuilder builder = new DocumentBuilder(clonedDoc);
            // Move the cursor to the end of the document before inserting the field.
            builder.MoveToDocumentEnd();

            // Insert the MERGEBARCODE field and configure its properties.
            FieldMergeBarcode barcodeField = (FieldMergeBarcode)builder.InsertField(FieldType.FieldMergeBarcode, true);
            barcodeField.BarcodeType = "QR";                 // QR code type.
            barcodeField.BarcodeValue = "MyQRCode";          // The value to encode.
            barcodeField.BackgroundColor = "0xF8BD69";       // Optional background colour.
            barcodeField.ForegroundColor = "0xB5413B";       // Optional foreground colour.
            barcodeField.ErrorCorrectionLevel = "3";         // High error correction.
            barcodeField.ScalingFactor = "250";              // Scale the symbol.
            barcodeField.SymbolHeight = "1000";              // Height in TWIPS.
            barcodeField.SymbolRotation = "0";               // No rotation.

            // Append another document to the cloned document.
            Document docToAppend = new Document("Append.docx");
            clonedDoc.AppendDocument(docToAppend, ImportFormatMode.KeepSourceFormatting);

            // Split the cloned document – extract the first page into a new document.
            Document splitDoc = clonedDoc.ExtractPages(0, 1);

            // Save the documents.
            originalDoc.Save("Original.docx");
            clonedDoc.Save("ClonedWithBarcode.docx");
            splitDoc.Save("FirstPageSplit.docx");
        }
    }
}
