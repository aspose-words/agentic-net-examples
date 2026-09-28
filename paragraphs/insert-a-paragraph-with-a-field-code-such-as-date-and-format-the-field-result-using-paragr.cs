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

        // Insert a new paragraph and add a DATE field.
        builder.Writeln();                     // Ensure we are on a new paragraph.
        builder.InsertField(FieldType.FieldDate, true);

        // Apply paragraph formatting to the paragraph that contains the field.
        Paragraph paragraph = builder.CurrentParagraph;
        paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Center; // Center the text.
        paragraph.ParagraphFormat.SpaceAfter = 12; // Add space after the paragraph (points).

        // Save the document.
        doc.Save("Output.docx");
    }
}
