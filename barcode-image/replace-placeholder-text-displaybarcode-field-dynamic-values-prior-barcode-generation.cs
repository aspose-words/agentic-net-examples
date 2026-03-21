using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Fields;

class ReplaceDisplayBarcodePlaceholder
{
    static void Main()
    {
        // Create a new document and insert DISPLAYBARCODE fields with placeholder values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First barcode field with placeholder "PLACEHOLDER".
        builder.InsertField("DISPLAYBARCODE \"PLACEHOLDER\"");

        // Add a paragraph break.
        builder.Writeln();

        // Second barcode field with placeholder "ANOTHER_PLACEHOLDER".
        builder.InsertField("DISPLAYBARCODE \"ANOTHER_PLACEHOLDER\"");

        // Define dynamic values that will replace the placeholders.
        var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "PLACEHOLDER", "ABC123" },
            { "ANOTHER_PLACEHOLDER", "DEF456" }
        };

        // Iterate through all fields in the document.
        foreach (Field field in doc.Range.Fields)
        {
            // Process only DISPLAYBARCODE fields.
            if (field.Type != FieldType.FieldDisplayBarcode)
                continue;

            // Cast to the specific field type.
            FieldDisplayBarcode barcodeField = (FieldDisplayBarcode)field;

            // Get the current barcode value (the placeholder).
            string currentValue = barcodeField.BarcodeValue;

            // Replace placeholder with actual value if a match is found.
            if (replacements.TryGetValue(currentValue, out string newValue))
            {
                barcodeField.BarcodeValue = newValue;
                barcodeField.IsDirty = true;
                barcodeField.Update();
            }
        }

        // Save the modified document.
        doc.Save("ResultWithDynamicBarcodes.docx");
    }
}
