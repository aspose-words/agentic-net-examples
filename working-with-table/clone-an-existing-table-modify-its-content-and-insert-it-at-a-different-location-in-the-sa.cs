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

        // Add some introductory text.
        builder.Writeln("Document start");

        // Build the original table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Original Cell 1");
        builder.InsertCell();
        builder.Write("Original Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Original Cell 3");
        builder.InsertCell();
        builder.Write("Original Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Add a placeholder paragraph where the cloned table will be inserted.
        builder.Writeln("Insert location");

        // Locate the original table in the document.
        Table originalTable = doc.GetChildNodes(NodeType.Table, true)
                                 .Cast<Table>()
                                 .FirstOrDefault();

        if (originalTable == null)
            throw new InvalidOperationException("Original table not found.");

        // Clone the original table (deep clone).
        Table clonedTable = (Table)originalTable.Clone(true);

        // Modify the cloned table's content: change text of the first cell.
        Cell firstCell = clonedTable.Rows[0].Cells[0];
        firstCell.Paragraphs.Clear(); // Remove existing paragraphs.
        Paragraph newParagraph = new Paragraph(doc);
        newParagraph.AppendChild(new Run(doc, "Cloned Table - Modified"));
        firstCell.AppendChild(newParagraph);

        // Find the placeholder paragraph.
        Paragraph placeholderParagraph = doc.GetChildNodes(NodeType.Paragraph, true)
                                            .Cast<Paragraph>()
                                            .FirstOrDefault(p => p.GetText().Contains("Insert location"));

        if (placeholderParagraph == null)
            throw new InvalidOperationException("Placeholder paragraph not found.");

        // Insert the cloned table after the placeholder paragraph.
        placeholderParagraph.ParentNode.InsertAfter(clonedTable, placeholderParagraph);

        // Save the final document.
        string outputPath = "ClonedTableOutput.docx";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
