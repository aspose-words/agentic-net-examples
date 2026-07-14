using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Define temporary file paths for the source documents and the final merged document.
        string tempDir = Path.Combine(Directory.GetCurrentDirectory(), "TempTables");
        Directory.CreateDirectory(tempDir);
        string srcPath1 = Path.Combine(tempDir, "SourceTable1.docx");
        string srcPath2 = Path.Combine(tempDir, "SourceTable2.docx");
        string resultPath = Path.Combine(tempDir, "MergedTables.docx");

        // -----------------------------------------------------------------
        // Create first source document with a simple 2x2 table.
        // -----------------------------------------------------------------
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);

        // Build table.
        builder1.StartTable();
        builder1.InsertCell();
        builder1.Write("Doc1 - Row1, Cell1");
        builder1.InsertCell();
        builder1.Write("Doc1 - Row1, Cell2");
        builder1.EndRow();

        builder1.InsertCell();
        builder1.Write("Doc1 - Row2, Cell1");
        builder1.InsertCell();
        builder1.Write("Doc1 - Row2, Cell2");
        builder1.EndRow();
        builder1.EndTable();

        // Save the first source document.
        srcDoc1.Save(srcPath1);

        // -----------------------------------------------------------------
        // Create second source document with a formatted 2x2 table.
        // -----------------------------------------------------------------
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);

        // Apply some cell shading to demonstrate formatting preservation.
        builder2.StartTable();

        // First row, first cell with light blue background.
        builder2.InsertCell();
        builder2.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
        builder2.Write("Doc2 - Row1, Cell1");

        // First row, second cell with light green background.
        builder2.InsertCell();
        builder2.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
        builder2.Write("Doc2 - Row1, Cell2");
        builder2.EndRow();

        // Second row, first cell with light pink background.
        builder2.InsertCell();
        builder2.CellFormat.Shading.BackgroundPatternColor = Color.LightPink;
        builder2.Write("Doc2 - Row2, Cell1");

        // Second row, second cell with light yellow background.
        builder2.InsertCell();
        builder2.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
        builder2.Write("Doc2 - Row2, Cell2");
        builder2.EndRow();

        builder2.EndTable();

        // Save the second source document.
        srcDoc2.Save(srcPath2);

        // -----------------------------------------------------------------
        // Create the destination document that will receive the tables.
        // -----------------------------------------------------------------
        Document dstDoc = new Document();

        // Load the first source document and import its table.
        Document loadDoc1 = new Document(srcPath1);
        Table table1 = loadDoc1.FirstSection.Body.Tables[0];
        NodeImporter importer1 = new NodeImporter(loadDoc1, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable1 = importer1.ImportNode(table1, true);
        dstDoc.FirstSection.Body.AppendChild(importedTable1);

        // Load the second source document and import its table.
        Document loadDoc2 = new Document(srcPath2);
        Table table2 = loadDoc2.FirstSection.Body.Tables[0];
        NodeImporter importer2 = new NodeImporter(loadDoc2, dstDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable2 = importer2.ImportNode(table2, true);
        dstDoc.FirstSection.Body.AppendChild(importedTable2);

        // Save the merged document.
        dstDoc.Save(resultPath);

        // Simple validation to ensure the output file was created.
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException("Merged document was not saved correctly.");
        }

        // Cleanup temporary source files (optional).
        // File.Delete(srcPath1);
        // File.Delete(srcPath2);
    }
}
