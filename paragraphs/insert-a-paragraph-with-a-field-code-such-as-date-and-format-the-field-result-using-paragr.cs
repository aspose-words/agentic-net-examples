using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure paragraph formatting: center alignment and a space after the paragraph.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.ParagraphFormat.SpaceAfter = 12; // points

        // Write some introductory text.
        builder.Write("Current date: ");

        // Insert a DATE field with a custom date format using the \\@ switch.
        // The field result will be formatted as "Wednesday, March 27, 2024".
        Field dateField = builder.InsertField("DATE \\@ \"dddd, MMMM dd, yyyy\"");

        // Finish the paragraph.
        builder.Writeln();

        // Ensure all fields are up‑to‑date before saving.
        doc.UpdateFields();

        // Save the document to the current directory.
        doc.Save("ParagraphWithDateField.docx");
    }
}
