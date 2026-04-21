using System;
using System.IO;
using System.Drawing; // For Color
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Folder for temporary files.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeTablesDemo");
        Directory.CreateDirectory(workDir);

        // Paths for the sample source documents.
        string sourcePath1 = Path.Combine(workDir, "Source1.docx");
        string sourcePath2 = Path.Combine(workDir, "Source2.docx");
        // Path for the merged output document.
        string mergedPath = Path.Combine(workDir, "Merged.docx");

        // -----------------------------------------------------------------
        // 1. Create two source documents, each containing a formatted table.
        // -----------------------------------------------------------------
        CreateSampleDocument(sourcePath1, "First", Color.LightBlue);
        CreateSampleDocument(sourcePath2, "Second", Color.LightGreen);

        // -----------------------------------------------------------------
        // 2. Create the destination document that will hold the merged tables.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Merged Tables:");
        destBuilder.Writeln(); // empty line before first table.

        // -----------------------------------------------------------------
        // 3. Load each source document, import its table, and append it.
        // -----------------------------------------------------------------
        foreach (string srcPath in new[] { sourcePath1, sourcePath2 })
        {
            Document srcDoc = new Document(srcPath);
            // Assume each source contains exactly one table.
            Table srcTable = srcDoc.FirstSection.Body.Tables[0];

            // Import the table node into the destination document preserving formatting.
            NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedTable = importer.ImportNode(srcTable, true);

            // Append the imported table to the destination body.
            destDoc.FirstSection.Body.AppendChild(importedTable);
            // Add a blank paragraph after each table for visual separation.
            destDoc.FirstSection.Body.AppendChild(new Paragraph(destDoc));
        }

        // -----------------------------------------------------------------
        // 4. Save the merged document.
        // -----------------------------------------------------------------
        destDoc.Save(mergedPath);

        // -----------------------------------------------------------------
        // 5. Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(mergedPath))
            throw new Exception("Merged document was not created.");

        Console.WriteLine($"Merged document saved to: {mergedPath}");
    }

    // Helper method to create a sample document with a single table.
    private static void CreateSampleDocument(string filePath, string tableLabel, Color shadingColor)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write($"{tableLabel} Table - Row 1, Cell 1");
        builder.InsertCell();
        builder.Write($"{tableLabel} Table - Row 1, Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write($"{tableLabel} Table - Row 2, Cell 1");
        builder.InsertCell();
        builder.Write($"{tableLabel} Table - Row 2, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Apply simple shading to the whole table to demonstrate formatting preservation.
        // Use SetShading because Table.Shading property does not exist.
        table.SetShading(TextureIndex.TextureSolid, shadingColor, Color.Empty);

        // Save the sample document.
        doc.Save(filePath);
    }
}
