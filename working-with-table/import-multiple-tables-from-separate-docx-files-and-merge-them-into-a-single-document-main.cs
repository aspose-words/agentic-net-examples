using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ---------- Create first source document with a simple table ----------
        Document sourceDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(sourceDoc1);
        builder1.StartTable();

        // Row 1
        builder1.InsertCell();
        builder1.Write("Doc1 - Cell A1");
        builder1.InsertCell();
        builder1.Write("Doc1 - Cell A2");
        builder1.EndRow();

        // Row 2
        builder1.InsertCell();
        builder1.Write("Doc1 - Cell B1");
        builder1.InsertCell();
        builder1.Write("Doc1 - Cell B2");
        builder1.EndRow();

        builder1.EndTable();
        string sourcePath1 = "Source1.docx";
        sourceDoc1.Save(sourcePath1);

        // ---------- Create second source document with a different table ----------
        Document sourceDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(sourceDoc2);
        builder2.StartTable();

        // Row 1
        builder2.InsertCell();
        builder2.Write("Doc2 - Cell X1");
        builder2.InsertCell();
        builder2.Write("Doc2 - Cell X2");
        builder2.EndRow();

        // Row 2
        builder2.InsertCell();
        builder2.Write("Doc2 - Cell Y1");
        builder2.InsertCell();
        builder2.Write("Doc2 - Cell Y2");
        builder2.EndRow();

        builder2.EndTable();
        string sourcePath2 = "Source2.docx";
        sourceDoc2.Save(sourcePath2);

        // ---------- Destination document that will contain merged tables ----------
        Document destDoc = new Document();

        // Array of source file paths to import.
        string[] sourceFiles = { sourcePath1, sourcePath2 };

        foreach (string srcPath in sourceFiles)
        {
            // Load the source document.
            Document srcDoc = new Document(srcPath);

            // Find the first table in the source document.
            NodeCollection tables = srcDoc.GetChildNodes(NodeType.Table, true);
            if (tables.Count == 0)
                throw new InvalidOperationException($"No table found in {srcPath}.");

            Table srcTable = (Table)tables[0];

            // Import the table into the destination document, preserving formatting.
            NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedNode = importer.ImportNode(srcTable, true);

            // Append the imported table to the destination document's body.
            destDoc.FirstSection.Body.AppendChild(importedNode);
        }

        // Save the merged document.
        string outputPath = "Merged.docx";
        destDoc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The merged document was not created.", outputPath);
    }
}
