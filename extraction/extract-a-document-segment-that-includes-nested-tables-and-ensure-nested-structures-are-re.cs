using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractNestedTablesExample
{
    public static void Main()
    {
        // Create a source document that contains a nested table structure.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before the tables.");

        // ---------- Outer table ----------
        Table outerTable = builder.StartTable();

        // First row of the outer table.
        builder.InsertCell();
        builder.Writeln("Outer Cell 1");
        builder.InsertCell();
        builder.Writeln("Outer Cell 2");
        builder.EndRow();

        // Second row of the outer table – this cell will contain the inner table.
        builder.InsertCell();

        // ---------- Inner table (nested) ----------
        Table innerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Inner Cell 1");
        builder.InsertCell();
        builder.Writeln("Inner Cell 2");
        builder.EndRow();
        builder.EndTable(); // End of inner table.

        // Continue the outer table's second row.
        builder.InsertCell();
        builder.Writeln("Outer Cell 3");
        builder.EndRow();

        builder.EndTable(); // End of outer table.

        builder.Writeln("Paragraph after the tables.");

        // -----------------------------------------------------------------
        // Extract the outer table (which includes the nested inner table) into a new document.
        // -----------------------------------------------------------------

        // Ensure the outer table was created.
        if (outerTable == null)
            throw new InvalidOperationException("Failed to create the outer table in the source document.");

        // Prepare the destination document.
        Document destDoc = new Document();
        destDoc.RemoveAllChildren(); // Remove the default empty section/paragraph.
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Import the outer table (with its nested content) into the destination document.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(outerTable, true);
        destBody.AppendChild(importedTable);

        // Save the extracted segment to a file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedNestedTables.docx");
        destDoc.Save(outputPath);

        // Validate that the output file was created and contains at least one table.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The extracted document was not saved.", outputPath);

        Document verificationDoc = new Document(outputPath);
        int tableCount = verificationDoc.GetChildNodes(NodeType.Table, true).Count;
        if (tableCount == 0)
            throw new InvalidOperationException("The extracted document does not contain any tables.");

        // Inform the user (console output is allowed for informational purposes).
        Console.WriteLine($"Extraction complete. Output saved to: {outputPath}");
        Console.WriteLine($"Extracted document contains {tableCount} table(s).");
    }
}
