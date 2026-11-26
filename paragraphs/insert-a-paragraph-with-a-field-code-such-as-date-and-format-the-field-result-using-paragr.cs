using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph.
        builder.Writeln(); // Moves cursor to a new paragraph.

        // Insert a DATE field. The field will be updated automatically.
        builder.InsertField(FieldType.FieldDate, true);

        // Apply paragraph formatting to the paragraph that contains the field.
        // For example, center the paragraph and add some spacing.
        Paragraph paragraph = builder.CurrentParagraph;
        paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        paragraph.ParagraphFormat.SpaceAfter = 12; // 12 points after spacing.

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
