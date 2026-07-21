using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Replacing;

namespace AsposeWordsTableReplace
{
    public class Program
    {
        public static void Main()
        {
            // Folder for generated files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a source document that contains an original table.
            // -----------------------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

            // Build a simple 2x2 table with placeholder text "Old".
            srcBuilder.StartTable();
            srcBuilder.InsertCell();
            srcBuilder.Write("Old Row 1, Cell 1");
            srcBuilder.InsertCell();
            srcBuilder.Write("Old Row 1, Cell 2");
            srcBuilder.EndRow();

            srcBuilder.InsertCell();
            srcBuilder.Write("Old Row 2, Cell 1");
            srcBuilder.InsertCell();
            srcBuilder.Write("Old Row 2, Cell 2");
            srcBuilder.EndRow();
            srcBuilder.EndTable();

            string sourcePath = Path.Combine(outputDir, "Original.docx");
            sourceDoc.Save(sourcePath);

            // -----------------------------------------------------------------
            // 2. Load the source document and create a new table that will replace the old one.
            // -----------------------------------------------------------------
            Document targetDoc = new Document(sourcePath);

            // Build the replacement table in a separate temporary document.
            Document templateDoc = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);

            tmplBuilder.StartTable();
            tmplBuilder.InsertCell();
            tmplBuilder.Write("New Row 1, Cell 1");
            tmplBuilder.InsertCell();
            tmplBuilder.Write("New Row 1, Cell 2");
            tmplBuilder.EndRow();

            tmplBuilder.InsertCell();
            tmplBuilder.Write("New Row 2, Cell 1");
            tmplBuilder.InsertCell();
            tmplBuilder.Write("New Row 2, Cell 2");
            tmplBuilder.EndRow();
            tmplBuilder.EndTable();

            // The table we just built.
            Table templateTable = (Table)templateDoc.GetChild(NodeType.Table, 0, true);

            // Import the template table into the target document.
            NodeImporter importer = new NodeImporter(templateDoc, targetDoc, ImportFormatMode.KeepSourceFormatting);
            Table newTable = (Table)importer.ImportNode(templateTable, true);

            // -----------------------------------------------------------------
            // 3. Locate the original table node in the document.
            // -----------------------------------------------------------------
            Table originalTable = (Table)targetDoc.GetChild(NodeType.Table, 0, true);
            if (originalTable == null)
                throw new InvalidOperationException("Original table not found.");

            // -----------------------------------------------------------------
            // 4. Replace the original table with the new one.
            // -----------------------------------------------------------------
            CompositeNode parent = originalTable.ParentNode as CompositeNode;
            if (parent == null)
                throw new InvalidOperationException("Original table does not have a valid parent.");

            // Insert the new table after the original table.
            parent.InsertAfter(newTable, originalTable);
            // Remove the original table.
            originalTable.Remove();

            // -----------------------------------------------------------------
            // 5. Save the modified document.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(outputDir, "Result.docx");
            targetDoc.Save(resultPath);

            // Simple validation that the file was created.
            if (!File.Exists(resultPath))
                throw new FileNotFoundException("Result document was not saved.", resultPath);
        }
    }
}
