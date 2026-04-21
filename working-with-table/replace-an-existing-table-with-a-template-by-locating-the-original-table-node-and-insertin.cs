using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableReplaceExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a source document that contains the original table.
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
            srcBuilder.StartTable();
            srcBuilder.InsertCell();
            srcBuilder.Write("Old Cell 1");
            srcBuilder.InsertCell();
            srcBuilder.Write("Old Cell 2");
            srcBuilder.EndRow();
            srcBuilder.EndTable();

            // Save the source document (optional, just to illustrate the initial file).
            string sourcePath = "Source.docx";
            sourceDoc.Save(sourcePath);

            // Create a template document that contains the replacement table.
            Document templateDoc = new Document();
            DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
            tmplBuilder.StartTable();
            tmplBuilder.InsertCell();
            tmplBuilder.Write("New Cell A");
            tmplBuilder.InsertCell();
            tmplBuilder.Write("New Cell B");
            tmplBuilder.EndRow();
            tmplBuilder.EndTable();

            // Retrieve the template table from the template document.
            Table templateTable = templateDoc.FirstSection.Body.Tables[0];

            // Locate the original table in the source document.
            Table oldTable = sourceDoc.FirstSection.Body.Tables[0];

            // Import the template table into the source document.
            NodeImporter importer = new NodeImporter(templateDoc, sourceDoc, ImportFormatMode.KeepSourceFormatting);
            Table importedTable = (Table)importer.ImportNode(templateTable, true);

            // Insert the imported table after the old table and then remove the old table.
            oldTable.ParentNode.InsertAfter(importedTable, oldTable);
            oldTable.Remove();

            // Save the modified document.
            string resultPath = "Result.docx";
            sourceDoc.Save(resultPath);

            // Simple validation to ensure the output file was created.
            if (!File.Exists(resultPath))
                throw new InvalidOperationException("The result document was not saved correctly.");
        }
    }
}
