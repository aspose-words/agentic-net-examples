using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a styled paragraph (Heading 1) and some text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Table inserted below:");

        // Retrieve the paragraph we just added.
        Paragraph styledParagraph = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true).Last();

        // Move the builder to the styled paragraph.
        builder.MoveTo(styledParagraph);

        // Optionally add a line break so the table starts on a new line.
        builder.Writeln();

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableInStyledParagraph.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Indicate successful completion.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
