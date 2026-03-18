using System;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply paragraph formatting that will affect the field result.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center; // Center the paragraph.
        builder.Font.Size = 14;                                        // Set font size.
        builder.Font.Bold = true;                                      // Make the text bold.

        // Insert a DATE field and update it immediately.
        // This uses the InsertField overload that takes a FieldType, complying with the provided API.
        Field dateField = builder.InsertField(FieldType.FieldDate, true);

        // Optionally, specify a custom date format using the field's Format property.
        // The format corresponds to the \@ switch of a DATE field.
        dateField.Format.DateTimeFormat = "dddd, MMMM dd, yyyy";
        dateField.Update(); // Re‑calculate the field to apply the new format.

        // Save the document to disk.
        doc.Save("ParagraphWithDateField.docx");
    }
}
