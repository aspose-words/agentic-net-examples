using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableReplace
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
            Directory.CreateDirectory(artifactsDir);

            // -----------------------------------------------------------------
            // 1. Create a source document that contains an original table.
            // -----------------------------------------------------------------
            string originalPath = Path.Combine(artifactsDir, "Original.docx");
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

            srcBuilder.Writeln("Source document with the original table:");
            srcBuilder.StartTable();
            srcBuilder.InsertCell();
            srcBuilder.Write("Original Cell 1");
            srcBuilder.InsertCell();
            srcBuilder.Write("Original Cell 2");
            srcBuilder.EndRow();
            srcBuilder.EndTable();

            sourceDoc.Save(originalPath);

            // -----------------------------------------------------------------
            // 2. Create a template document that contains the replacement table.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(artifactsDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);

            tmplBuilder.Writeln("Template document with the new table:");
            tmplBuilder.StartTable();
            tmplBuilder.InsertCell();
            tmplBuilder.Write("New Cell A");
            tmplBuilder.InsertCell();
            tmplBuilder.Write("New Cell B");
            tmplBuilder.EndRow();
            tmplBuilder.EndTable();

            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the source document and locate the original table node.
            // -----------------------------------------------------------------
            Document targetDoc = new Document(originalPath);
            Table originalTable = (Table)targetDoc.GetChild(NodeType.Table, 0, true);
            if (originalTable == null)
                throw new InvalidOperationException("Original table not found in the source document.");

            // -----------------------------------------------------------------
            // 4. Load the template table and import it into the target document.
            // -----------------------------------------------------------------
            Document tmplDocForImport = new Document(templatePath);
            Table templateTable = (Table)tmplDocForImport.GetChild(NodeType.Table, 0, true);
            if (templateTable == null)
                throw new InvalidOperationException("Template table not found in the template document.");

            NodeImporter importer = new NodeImporter(tmplDocForImport, targetDoc, ImportFormatMode.KeepSourceFormatting);
            Table importedTable = (Table)importer.ImportNode(templateTable, true);

            // -----------------------------------------------------------------
            // 5. Replace the original table with the imported template table.
            // -----------------------------------------------------------------
            // Insert the new table after the original one, then remove the original.
            originalTable.ParentNode.InsertAfter(importedTable, originalTable);
            originalTable.Remove();

            // -----------------------------------------------------------------
            // 6. Save the resulting document.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(artifactsDir, "Result.docx");
            targetDoc.Save(resultPath);

            // Verify that the output file was created.
            if (!File.Exists(resultPath))
                throw new FileNotFoundException("Result document was not saved correctly.", resultPath);
        }
    }
}
