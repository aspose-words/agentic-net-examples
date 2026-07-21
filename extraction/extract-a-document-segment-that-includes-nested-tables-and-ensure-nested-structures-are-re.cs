using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------
        // Create a sample source document
        // -----------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Intro paragraph before the table.
        builder.Writeln("Intro paragraph before the table.");

        // Start the outer table.
        builder.StartTable();

        // First cell of the outer table.
        builder.InsertCell();
        builder.Writeln("Outer cell 1");

        // Capture the cell that currently contains the paragraph we just wrote.
        Cell outerCell = builder.CurrentParagraph.ParentNode as Cell;
        if (outerCell == null)
            throw new InvalidOperationException("Failed to obtain the outer table cell.");

        // Move the cursor into the first paragraph of that cell to insert the inner table.
        builder.MoveTo(outerCell.FirstParagraph);

        // Start the inner (nested) table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Inner cell 1");
        builder.InsertCell();
        builder.Writeln("Inner cell 2");
        builder.EndRow();
        builder.EndTable();

        // Return to the outer table row and finish the first row.
        builder.MoveTo(outerCell.FirstParagraph);
        builder.EndRow();

        // Second cell of the outer table (same row).
        builder.InsertCell();
        builder.Writeln("Outer cell 2");
        builder.EndRow();

        // End the outer table.
        builder.EndTable();

        // Paragraph after the table.
        builder.Writeln("Paragraph after the table.");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------
        // Load the document for extraction
        // -----------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the first table (which contains the nested table).
        Table outerTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (outerTable == null)
            throw new InvalidOperationException("No table found in the source document.");

        // Clone the table, preserving all nested structures.
        Table clonedTable = outerTable.Clone(true) as Table;
        if (clonedTable == null)
            throw new InvalidOperationException("Failed to clone the table.");

        // -----------------------------
        // Build a new document to hold the extracted segment
        // -----------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Import the cloned table into the new document (required because it was created from a different document).
        Node importedTable = extractedDoc.ImportNode(clonedTable, true);
        body.AppendChild(importedTable);

        // Save the extracted segment.
        const string extractedPath = "extracted.docx";
        extractedDoc.Save(extractedPath);

        // Verify that the output file was created.
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("The extracted document was not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
