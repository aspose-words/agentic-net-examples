using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ------------------------------------------------------------
        // Set paragraph formatting that will be applied to the next
        // paragraph (the one that will contain the DATE field).
        // ------------------------------------------------------------
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center; // Center the text.
        builder.ParagraphFormat.SpaceAfter = 12;                       // Add 12 points spacing after the paragraph.

        // Write some introductory text and start a new paragraph.
        builder.Writeln("Current date:");

        // Insert a DATE field. The field is updated immediately (second argument = true).
        Field dateField = builder.InsertField(FieldType.FieldDate, true);

        // Apply a custom date/time format to the field result via the FieldFormat object.
        // This corresponds to the \\@ switch in a Word field.
        dateField.Format.DateTimeFormat = "dddd, MMMM dd, yyyy";

        // Ensure all fields in the document are up‑to‑date before saving.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("ParagraphWithDateField.docx");
    }
}
