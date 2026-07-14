using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph formatting that will be applied to the next paragraph.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;   // Center the paragraph.
        builder.ParagraphFormat.SpaceAfter = 12;                        // Add space after the paragraph.

        // Start a new paragraph and write some introductory text.
        builder.Writeln("Current date:");

        // Write the label and insert a DATE field on the same line.
        builder.Write("Date: ");
        Field dateField = builder.InsertField(FieldType.FieldDate, true); // Insert and update the field.

        // Apply a custom date format to the field result (e.g., "Monday, January 01, 2024").
        dateField.Format.DateTimeFormat = "dddd, MMMM dd, yyyy";

        // End the paragraph.
        builder.Writeln();

        // Ensure all fields are up‑to‑date before saving.
        doc.UpdateFields();

        // Save the document to the output file.
        doc.Save("Output.docx");
    }
}
