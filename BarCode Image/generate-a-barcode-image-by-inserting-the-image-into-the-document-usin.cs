using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class GenerateBarcodeDocument
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example barcode data.
        string barcodeType = "QR";          // Supported types: QR, CODE_128, etc.
        string barcodeValue = "ABC123";    // The value to encode.

        // Insert a BARCODE field. The syntax is: { BARCODE <type> "<value>" }
        // Aspose.Words will render the barcode when the document is saved.
        builder.InsertField($"BARCODE {barcodeType} \"{barcodeValue}\"");

        // Save the document as DOCX.
        doc.Save("BarcodeDocument.docx");
    }
}
