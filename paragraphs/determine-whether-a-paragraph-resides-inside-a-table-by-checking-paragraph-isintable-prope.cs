using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Paragraph that is NOT inside a table.
        builder.Writeln("Paragraph outside any table.");

        // Build a simple 1‑row, 2‑cell table.
        Table table = builder.StartTable();

        // First cell with a paragraph.
        builder.InsertCell();
        builder.Writeln("Paragraph inside first cell.");

        // Second cell with a paragraph.
        builder.InsertCell();
        builder.Writeln("Paragraph inside second cell.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Another paragraph after the table.
        builder.Writeln("Another paragraph outside any table.");

        // Iterate through all paragraphs in the document.
        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in allParagraphs)
        {
            // Determine whether the paragraph is an immediate child of a Cell.
            bool isInCell = para.IsInCell;

            // Get the paragraph text without the trailing paragraph break.
            string text = para.GetText().Trim();

            // Output the result.
            Console.WriteLine($"\"{text}\"  IsInCell = {isInCell}");
        }

        // Save the document to the local file system.
        doc.Save("ParagraphInTable.docx");
    }
}
