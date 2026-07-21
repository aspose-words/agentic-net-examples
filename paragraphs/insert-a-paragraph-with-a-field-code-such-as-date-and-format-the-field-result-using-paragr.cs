using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsParagraphFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // DocumentBuilder is used to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a heading for the paragraph.
            builder.Writeln("Report generated on:");

            // Insert a DATE field with a custom format.
            // The field code includes the \\@ switch that defines the date format.
            Field dateField = builder.InsertField("DATE \\@ \"dddd, MMMM dd, yyyy\"");

            // Ensure all fields are up‑to‑date (optional, as InsertField updates the field immediately).
            doc.UpdateFields();

            // Apply paragraph formatting to the paragraph that now contains the DATE field.
            Paragraph currentParagraph = builder.CurrentParagraph;
            currentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            currentParagraph.ParagraphFormat.SpaceAfter = 12; // points

            // Save the document.
            string outputPath = "ParagraphWithDateField.docx";
            doc.Save(outputPath);
        }
    }
}
