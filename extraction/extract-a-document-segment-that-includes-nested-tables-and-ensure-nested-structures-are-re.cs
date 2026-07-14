using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------
        // 1. Create a source document that contains a nested table.
        // -----------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Intro paragraph.
        builder.Writeln("Intro paragraph.");

        // Begin outer table.
        builder.StartTable();

        // First cell of outer table.
        builder.InsertCell();
        builder.Writeln("Outer cell 1");

        // Capture the cell that we are currently in.
        Cell outerCell = builder.CurrentParagraph.ParentNode as Cell;
        if (outerCell == null)
            throw new InvalidOperationException("Failed to retrieve the outer cell.");

        // Move the cursor to the first paragraph of the outer cell and create an inner table inside it.
        builder.MoveTo(outerCell.FirstParagraph);
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Inner cell 1");
        builder.EndTable(); // End inner table.

        // Return to the outer cell and finish the first row.
        builder.MoveTo(outerCell.FirstParagraph);
        builder.EndRow();

        // Second cell of outer table.
        builder.InsertCell();
        builder.Writeln("Outer cell 2");
        builder.EndRow(); // End the second row.

        // End outer table.
        builder.EndTable();

        // Paragraph after the table.
        builder.Writeln("After paragraph.");

        // Save the source document to a local file.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------
        // 2. Load the source document from disk.
        // -----------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the first table in the document (the outer table).
        Table outerTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (outerTable == null)
            throw new InvalidOperationException("No table found in the source document.");

        // -----------------------------
        // 3. Create a new empty document that will hold the extracted segment.
        // -----------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Ensure the document is completely empty.

        // Build the minimal required structure: Section -> Body.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // -----------------------------
        // 4. Import the outer table (including its nested inner table) into the new document.
        //    Use ImportNode to copy nodes from a different document safely.
        // -----------------------------
        Node importedTable = resultDoc.ImportNode(outerTable, true);
        resultBody.AppendChild(importedTable);

        // -----------------------------
        // 5. Save the extracted segment.
        // -----------------------------
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // The program finishes silently; success can be inferred from the absence of exceptions.
    }
}
