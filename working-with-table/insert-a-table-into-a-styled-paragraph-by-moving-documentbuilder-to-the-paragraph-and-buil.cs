using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder attached to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph and apply a style to it (e.g., Heading1).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Styled Paragraph");

        // Retrieve the paragraph we just created.
        Paragraph styledParagraph = doc.FirstSection.Body.Paragraphs[0];

        // Move the builder's cursor to the styled paragraph.
        builder.MoveTo(styledParagraph);

        // Build a 2x2 table at the current cursor position.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Save the document.
        string outputPath = "TableInStyledParagraph.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }
}
