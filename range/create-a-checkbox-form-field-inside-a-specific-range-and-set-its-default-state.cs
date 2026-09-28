using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

namespace AsposeWordsRangeCheckboxExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some initial content.
            builder.Writeln("First paragraph.");
            builder.Writeln("Second paragraph where the checkbox will be placed:");

            // Locate the second paragraph (index 1) to define the target range.
            Paragraph targetParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);

            // Move the builder to the start of the target paragraph's range.
            builder.MoveTo(targetParagraph);
            // Write preceding text inside the same range.
            builder.Write("Accept terms: ");

            // Insert a checkbox form field with a default checked state.
            // Parameters: field name, default state (true = checked), size in points.
            builder.InsertCheckBox("TermsCheckBox", true, 10);

            // Optionally add a line break after the checkbox.
            builder.Writeln();

            // Save the document to disk.
            doc.Save("Output.docx");
        }
    }
}
