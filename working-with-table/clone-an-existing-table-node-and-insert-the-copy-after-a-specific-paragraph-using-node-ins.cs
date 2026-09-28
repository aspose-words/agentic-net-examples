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

        // Add a paragraph before the table.
        builder.Writeln("Paragraph before table.");

        // Build a simple 2x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Add another paragraph after the table.
        builder.Writeln("Paragraph after table.");

        // Locate the first paragraph (the one before the table).
        Paragraph firstParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        if (firstParagraph == null)
            throw new InvalidOperationException("First paragraph not found.");

        // Locate the first table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("Original table not found.");

        // Clone the table (deep clone).
        Node clonedTable = originalTable.Clone(true);

        // Insert the cloned table after the first paragraph.
        firstParagraph.ParentNode.InsertAfter(clonedTable, firstParagraph);

        // Save the document.
        string outputPath = "ClonedTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);
    }
}
