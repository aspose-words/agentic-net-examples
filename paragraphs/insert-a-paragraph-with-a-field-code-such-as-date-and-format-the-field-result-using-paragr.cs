using System;
using Aspose.Words;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set paragraph formatting that will affect the field result.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.ParagraphFormat.SpaceAfter = 12; // points

        // Insert a DATE field; the field result will inherit the paragraph formatting.
        builder.InsertField(FieldType.FieldDate, true);

        // Add a line break after the field for clarity.
        builder.Writeln();

        // Save the document.
        doc.Save("Output.docx");
    }
}
