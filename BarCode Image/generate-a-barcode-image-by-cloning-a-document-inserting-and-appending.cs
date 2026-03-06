using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsBarcodeDemo
{
    class Program
    {
        static void Main()
        {
            // Load the main DOCX document.
            Document mainDoc = new Document("Source.docx");
            DocumentBuilder builder = new DocumentBuilder(mainDoc);

            // Insert a BARCODE field that will generate a US ZIP code barcode.
            // The field will be displayed as an image after updating fields.
            FieldBarcode barcodeField = (FieldBarcode)builder.InsertField(FieldType.FieldBarcode, true);
            barcodeField.PostalAddress = "12345";          // ZIP code to encode.
            barcodeField.IsUSPostalAddress = true;        // Specify US postal address.
            barcodeField.FacingIdentificationMark = "C";  // Optional FIM character.
            builder.Writeln(); // Move to the next line.

            // Clone the document (deep copy).
            Document clonedDoc = (Document)mainDoc.Clone();

            // Insert another document at the current cursor position.
            Document insertDoc = new Document("Insert.docx");
            builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

            // Append a third document to the end of the main document.
            Document appendDoc = new Document("Append.docx");
            mainDoc.AppendDocument(appendDoc, ImportFormatMode.KeepSourceFormatting);

            // Update all fields so that the BARCODE field renders the barcode image.
            mainDoc.UpdateFields();

            // Save the combined document.
            mainDoc.Save("CombinedResult.docx");

            // Save the cloned document (it contains the same barcode field).
            clonedDoc.Save("ClonedResult.docx");

            // Split the combined document into separate pages.
            // Each page is extracted into a new Document and saved individually.
            for (int pageIndex = 1; pageIndex <= mainDoc.PageCount; pageIndex++)
            {
                // Extract a single page (pages are 1‑based).
                Document pageDoc = mainDoc.ExtractPages(pageIndex, pageIndex);
                string pageFileName = $"Page_{pageIndex}.docx";
                pageDoc.Save(pageFileName);
            }
        }
    }
}
