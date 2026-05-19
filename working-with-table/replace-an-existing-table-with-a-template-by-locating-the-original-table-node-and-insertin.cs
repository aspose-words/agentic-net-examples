using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a source document that contains an original table.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // Build a simple 2x2 table.
        Table originalTable = srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("A1");
        srcBuilder.InsertCell();
        srcBuilder.Write("A2");
        srcBuilder.EndRow();

        srcBuilder.InsertCell();
        srcBuilder.Write("B1");
        srcBuilder.InsertCell();
        srcBuilder.Write("B2");
        srcBuilder.EndRow();
        srcBuilder.EndTable();

        // Save the source document (required by the rules).
        const string sourcePath = "Original.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document (demonstrates the load rule).
        // -----------------------------------------------------------------
        Document mainDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Create a template table in a separate temporary document.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);

        // Build a 3x3 table that will replace the original one.
        Table templateTable = tmplBuilder.StartTable();

        // Row 1
        tmplBuilder.InsertCell();
        tmplBuilder.Write("T1");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("T2");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("T3");
        tmplBuilder.EndRow();

        // Row 2
        tmplBuilder.InsertCell();
        tmplBuilder.Write("T4");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("T5");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("T6");
        tmplBuilder.EndRow();

        // Row 3
        tmplBuilder.InsertCell();
        tmplBuilder.Write("T7");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("T8");
        tmplBuilder.InsertCell();
        tmplBuilder.Write("T9");
        tmplBuilder.EndRow();

        tmplBuilder.EndTable();

        // -----------------------------------------------------------------
        // 4. Import the template table into the main document.
        // -----------------------------------------------------------------
        NodeImporter importer = new NodeImporter(templateDoc, mainDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(templateTable, true);

        // -----------------------------------------------------------------
        // 5. Locate the original table in the main document.
        // -----------------------------------------------------------------
        Node originalNode = mainDoc.GetChild(NodeType.Table, 0, true);
        if (originalNode == null)
            throw new InvalidOperationException("Original table not found.");

        // -----------------------------------------------------------------
        // 6. Replace the original table with the imported template table.
        // -----------------------------------------------------------------
        // Insert the new table after the original one.
        originalNode.ParentNode.InsertAfter(importedTable, originalNode);
        // Remove the original table.
        originalNode.Remove();

        // -----------------------------------------------------------------
        // 7. Save the resulting document.
        // -----------------------------------------------------------------
        const string resultPath = "Result.docx";
        mainDoc.Save(resultPath);

        // Verify that the file was created.
        if (!File.Exists(resultPath))
            throw new Exception("Result document was not saved correctly.");
    }
}
