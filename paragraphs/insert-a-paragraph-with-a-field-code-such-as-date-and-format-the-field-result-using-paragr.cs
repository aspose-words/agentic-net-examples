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

        // Insert a new empty paragraph where the field will be placed.
        builder.InsertParagraph();

        // Apply paragraph formatting: center alignment and a space after the paragraph.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.ParagraphFormat.SpaceAfter = 12; // points

        // Insert a DATE field with a custom date format.
        // The field code does not include the surrounding braces.
        builder.InsertField("DATE \\@ \"dddd, MMMM dd, yyyy\"");

        // Update all fields so the DATE field shows the current date.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("ParagraphWithDateField.docx");
    }
}
