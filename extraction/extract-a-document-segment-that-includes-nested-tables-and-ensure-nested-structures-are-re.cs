using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------
        // 1. Create a sample source DOCX with a nested table.
        // -----------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Intro paragraph.
        builder.Writeln("Intro paragraph before the tables.");

        // Outer table.
        builder.StartTable();

        // Row 1, Cell 1.
        builder.InsertCell();
        builder.Write("Outer Cell 1");

        // Row 1, Cell 2 – will contain a nested table.
        builder.InsertCell();

        // Nested table inside the current cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Nested A1");
        builder.InsertCell();
        builder.Write("Nested B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Nested A2");
        builder.InsertCell();
        builder.Write("Nested B2");
        builder.EndRow();
        builder.EndTable(); // End nested table.

        // End first row of outer table.
        builder.EndRow();

        // Row 2 of outer table.
        builder.InsertCell();
        builder.Write("Outer Cell 2");
        builder.InsertCell();
        builder.Write("Outer Cell 3");
        builder.EndRow();

        builder.EndTable(); // End outer table.

        // Paragraph after the tables.
        builder.Writeln("Paragraph after the tables.");

        // Save the source document to a deterministic local file.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------
        // 2. Load the source document and locate the outer table.
        // -----------------------------
        Document loadedDoc = new Document(sourcePath);
        NodeCollection tableNodes = loadedDoc.GetChildNodes(NodeType.Table, true);
        if (tableNodes.Count == 0)
            throw new InvalidOperationException("No tables were found in the source document.");

        Table outerTable = tableNodes[0] as Table;
        if (outerTable == null)
            throw new InvalidOperationException("Outer table not found in the source document.");

        // -----------------------------
        // 3. Prepare a fresh destination document.
        // -----------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // -----------------------------
        // 4. Import (clone) the outer table into the new document.
        // -----------------------------
        NodeImporter importer = new NodeImporter(loadedDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedNode = importer.ImportNode(outerTable, true);
        Table importedTable = importedNode as Table;
        if (importedTable == null)
            throw new InvalidOperationException("Failed to import the outer table.");

        body.AppendChild(importedTable);

        // -----------------------------
        // 5. Save the extracted document.
        // -----------------------------
        const string extractedPath = "extracted.docx";
        extractedDoc.Save(extractedPath);

        // Verify that the file was created.
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // -----------------------------
        // 6. Validate that the nested table is retained.
        // -----------------------------
        Document validationDoc = new Document(extractedPath);
        NodeCollection extractedTables = validationDoc.GetChildNodes(NodeType.Table, true);
        // Expect at least two tables: the outer table and its nested table.
        if (extractedTables.Count < 2)
            throw new InvalidOperationException("Nested table was not retained in the extracted document.");

        // -----------------------------
        // 7. Completion message.
        // -----------------------------
        Console.WriteLine("Extraction completed successfully. Files created: " + sourcePath + ", " + extractedPath);
    }
}
