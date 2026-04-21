using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with two sections, each containing a table with a long row.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Long text that will force the row to span multiple pages.
        string longText = string.Concat(Enumerable.Repeat("LongText ", 200));

        // First section.
        InsertTableWithLongRow(builder, longText);

        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section.
        InsertTableWithLongRow(builder, longText);

        // Save the source document.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // Split the document by sections, preserving complete table rows.
        int sectionIndex = 1;
        foreach (Section section in sourceDoc.Sections)
        {
            // Create a new empty document.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren();

            // Import the section into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(outputDir, $"Section_{sectionIndex}.docx");
            splitDoc.Save(splitPath);

            // Validate that the long row is intact.
            ValidateLongRow(splitPath, longText);

            // Verify file existence.
            if (!File.Exists(splitPath))
                throw new InvalidOperationException($"Split file not found: {splitPath}");

            sectionIndex++;
        }

        // Verify source file exists.
        if (!File.Exists(sourcePath))
            throw new InvalidOperationException($"Source file not found: {sourcePath}");

        // Indicate successful completion.
        Console.WriteLine("Document split completed successfully.");
    }

    private static void InsertTableWithLongRow(DocumentBuilder builder, string longText)
    {
        // Start a table.
        Table table = builder.StartTable();

        // Insert a single cell with the long text.
        builder.InsertCell();
        builder.Write(longText);
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Add a paragraph after the table to ensure proper layout.
        builder.Writeln();
    }

    private static void ValidateLongRow(string docPath, string expectedText)
    {
        Document doc = new Document(docPath);
        Table table = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().FirstOrDefault();
        if (table == null)
            throw new InvalidOperationException($"No table found in document: {docPath}");

        Row row = table.Rows[0];
        Cell cell = row.Cells[0];
        string cellText = cell.GetText().Replace("\a", "").Trim(); // Remove end-of-cell markers.

        if (!cellText.Contains(expectedText.Trim()))
            throw new InvalidOperationException($"Long row text was not preserved in document: {docPath}");
    }
}
