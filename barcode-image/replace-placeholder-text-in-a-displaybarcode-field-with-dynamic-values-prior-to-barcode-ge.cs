using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DISPLAYBARCODE field with a placeholder barcode value.
        FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        barcodeField.BarcodeType = "CODE128";
        barcodeField.BarcodeValue = "PLACEHOLDER";
        barcodeField.ForegroundColor = "0x000000";
        barcodeField.BackgroundColor = "0xFFFFFF";

        // Dynamic value that should replace the placeholder before barcode generation.
        string dynamicBarcodeValue = "987654321";

        // Replace the placeholder text in the field's BarcodeValue property.
        // Since the placeholder is stored in the field's property, we assign the new value directly.
        barcodeField.BarcodeValue = dynamicBarcodeValue;

        // Update fields to generate the barcode with the new value.
        doc.UpdateFields();

        // Save the document.
        doc.Save("output.docx");
    }
}
