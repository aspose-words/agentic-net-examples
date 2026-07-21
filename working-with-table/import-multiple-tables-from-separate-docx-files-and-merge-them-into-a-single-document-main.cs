using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Directory to store sample and result documents.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the source documents that contain individual tables.
        string sourcePath1 = Path.Combine(artifactsDir, "Table1.docx");
        string sourcePath2 = Path.Combine(artifactsDir, "Table2.docx");

        // Create first sample document with a simple 2x2 table.
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.StartTable();
        builder1.InsertCell();
        builder1.Write("Source 1 - Cell A1");
        builder1.InsertCell();
        builder1.Write("Source 1 - Cell A2");
        builder1.EndRow();
        builder1.InsertCell();
        builder1.Write("Source 1 - Cell B1");
        builder1.InsertCell();
        builder1.Write("Source 1 - Cell B2");
        builder1.EndRow();
        builder1.EndTable();
        srcDoc1.Save(sourcePath1);

        // Create second sample document with a 2x2 table that has shading to demonstrate formatting preservation.
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.StartTable();
        builder2.InsertCell();
        // Apply background shading to the first cell.
        builder2.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder2.Write("Source 2 - Cell A1 (shaded)");
        builder2.InsertCell();
        builder2.Write("Source 2 - Cell A2");
        builder2.EndRow();
        builder2.InsertCell();
        builder2.Write("Source 2 - Cell B1");
        builder2.InsertCell();
        builder2.Write("Source 2 - Cell B2");
        builder2.EndRow();
        builder2.EndTable();
        srcDoc2.Save(sourcePath2);

        // Destination document that will receive the imported tables.
        Document destDoc = new Document();

        // Array of source file paths to process.
        string[] sourceFiles = { sourcePath1, sourcePath2 };

        foreach (string srcPath in sourceFiles)
        {
            // Load the source document.
            Document srcDoc = new Document(srcPath);

            // Retrieve the first table from the source document.
            Table srcTable = srcDoc.FirstSection.Body.Tables[0];

            // Import the table into the destination document, preserving its original formatting.
            NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
            Table importedTable = (Table)importer.ImportNode(srcTable, true);

            // Insert a paragraph break before each imported table for visual separation (except before the first one).
            if (destDoc.FirstSection.Body.Tables.Count > 0)
            {
                destDoc.FirstSection.Body.AppendChild(new Paragraph(destDoc));
            }

            // Append the imported table to the destination document.
            destDoc.FirstSection.Body.AppendChild(importedTable);
        }

        // Save the merged document.
        string mergedPath = Path.Combine(artifactsDir, "MergedTables.docx");
        destDoc.Save(mergedPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(mergedPath))
        {
            throw new Exception("Merged document was not saved correctly.");
        }
    }
}
