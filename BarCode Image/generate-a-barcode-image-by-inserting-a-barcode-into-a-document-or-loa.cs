using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace BarcodeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to work with the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a MERGEBARCODE field with the desired barcode type and value.
            // The first argument is the barcode type (e.g., "CODE39").
            // The second argument is the data to encode.
            FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FieldMergeBarcode);
            fieldBuilder.AddArgument("CODE39");
            fieldBuilder.AddArgument("12345ABCDE");

            // Insert the field at the current position in the paragraph.
            Field field = fieldBuilder.BuildAndInsert(builder.CurrentParagraph);

            // Cast to FieldMergeBarcode to set additional properties, such as displaying the text.
            if (field is FieldMergeBarcode mergeBarcode)
            {
                mergeBarcode.DisplayText = true;          // Show the barcode data below the image.
                mergeBarcode.AddStartStopChar = true;     // Add start/stop characters for CODE39.
            }

            // Update all fields so the barcode image is generated.
            doc.UpdateFields();

            // Save the document containing the barcode.
            doc.Save("Barcode.docx");
        }
    }
}
