using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;

public sealed class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure there is a paragraph to host the field.
        builder.Writeln();

        // Insert a DISPLAYBARCODE field with a temporary placeholder value.
        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.InsertField(FieldType.FieldDisplayBarcode, true);
        field.BarcodeType = "CODE128";
        field.BarcodeValue = "PLACEHOLDER";
        field.ForegroundColor = "0x000000";
        field.BackgroundColor = "0xFFFFFF";

        // Replace the placeholder text with the actual barcode data.
        doc.Range.Replace("PLACEHOLDER", "1234567890", new FindReplaceOptions());

        // Update fields so the barcode is rendered with the new value.
        doc.UpdateFields();

        // Save the document.
        doc.Save("output.docx");
    }
}
