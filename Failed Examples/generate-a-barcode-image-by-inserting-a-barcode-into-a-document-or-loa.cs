// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a MERGEBARCODE field.
        // Argument 1 – barcode type (e.g., QR, CODE39, EAN13, etc.).
        // Argument 2 – barcode value to encode.
        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldMergeBarcode);
        fieldBuilder.AddArgument("QR");          // Barcode type.
        fieldBuilder.AddArgument("ABC123");      // Barcode value.

        // Insert the field at the end of the current paragraph.
        Field field = fieldBuilder.BuildAndInsert(builder.CurrentParagraph);

        // Cast to the specific field type to set additional properties.
        FieldMergeBarcode barcodeField = (FieldMergeBarcode)field;
        barcodeField.DisplayText = true;         // Show the encoded text under the barcode.
        barcodeField.ScalingFactor = 250;        // Scale the barcode (percentage).
        barcodeField.SymbolHeight = 1000;        // Height in TWIPS (1/1440 inch).

        // Update fields so the barcode image is generated.
        doc.UpdateFields();

        // Save the document containing the barcode.
        doc.Save("BarcodeDocument.docx");
    }
}
