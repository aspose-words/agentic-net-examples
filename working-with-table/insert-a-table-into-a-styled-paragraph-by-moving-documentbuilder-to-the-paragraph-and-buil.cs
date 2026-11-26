using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a styled paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Styled Paragraph");

        // Retrieve the paragraph we just added.
        Paragraph styledParagraph = doc.LastSection.Body.Paragraphs[doc.LastSection.Body.Paragraphs.Count - 1];

        // Move the builder to the styled paragraph.
        builder.MoveTo(styledParagraph);

        // Build a 2x2 table at the current builder position.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableInStyledParagraph.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
