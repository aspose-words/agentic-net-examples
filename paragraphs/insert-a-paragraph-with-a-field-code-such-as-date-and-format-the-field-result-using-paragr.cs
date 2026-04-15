using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for insertion and formatting.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply paragraph formatting that will affect the paragraph containing the field.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;   // Center the paragraph.
        builder.ParagraphFormat.SpaceAfter = 12;                         // Add space after the paragraph.

        // Write some introductory text.
        builder.Write("Current date: ");

        // Insert a DATE field and update it immediately.
        Field field = builder.InsertField(FieldType.FieldDate, true);

        // Apply a custom date/time format to the field result.
        field.Format.DateTimeFormat = "dddd, MMMM dd, yyyy";

        // Update the field to reflect the new format.
        field.Update();

        // End the paragraph.
        builder.Writeln();

        // Save the document to disk.
        doc.Save("ParagraphWithDateField.docx");
    }
}
