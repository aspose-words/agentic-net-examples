using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that is NOT inside a table.
        builder.Writeln("Paragraph outside any table.");

        // Build a simple 2‑cell table with paragraphs inside each cell.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Writeln("Paragraph inside first table cell.");

        // Second cell.
        builder.InsertCell();
        builder.Writeln("Paragraph inside second table cell.");

        // End the table.
        builder.EndTable();

        // Iterate through all paragraphs in the document and report whether they are inside a table cell.
        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in allParagraphs)
        {
            // Trim the paragraph text to remove the trailing paragraph break.
            string text = para.GetText().TrimEnd('\r', '\a');
            Console.WriteLine($"Text: \"{text}\" | IsInCell: {para.IsInCell}");
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ParagraphIsInCellDemo.docx");
        doc.Save(outputPath);
    }
}
