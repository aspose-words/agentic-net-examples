using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------
        // Create first sample document.
        // -----------------------------
        Document sourceDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(sourceDoc1);
        builder1.Writeln("Document 1");
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

        // ------------------------------
        // Create second sample document.
        // ------------------------------
        Document sourceDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(sourceDoc2);
        builder2.Writeln("Document 2");
        builder2.StartTable();
        builder2.InsertCell();
        builder2.Write("Doc2 - Row1, Cell1");
        builder2.InsertCell();
        builder2.Write("Doc2 - Row1, Cell2");
        builder2.InsertCell();
        builder2.Write("Doc2 - Row1, Cell3");
        builder2.EndRow();
        builder2.InsertCell();
        builder2.Write("Doc2 - Row2, Cell1");
        builder2.InsertCell();
        builder2.Write("Doc2 - Row2, Cell2");
        builder2.InsertCell();
        builder2.Write("Doc2 - Row2, Cell3");
        builder2.EndRow();
        builder2.EndTable();

        // Destination document that will hold all merged tables.
        Document destDoc = new Document();

        // Helper method to import all tables from a source document.
        void ImportTables(Document src)
        {
            // NodeImporter handles copying nodes between documents while preserving formatting.
            NodeImporter importer = new NodeImporter(src, destDoc, ImportFormatMode.KeepSourceFormatting);

            // Get all Table nodes from the source document.
            NodeCollection tables = src.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Import the table into the destination document.
                Node importedTable = importer.ImportNode(table, true);
                // Append the imported table to the body of the destination document.
                destDoc.FirstSection.Body.AppendChild(importedTable);
                // Add a paragraph break between tables for readability.
                destDoc.FirstSection.Body.AppendChild(new Paragraph(destDoc));
            }
        }

        // Import tables from both source documents.
        ImportTables(sourceDoc1);
        ImportTables(sourceDoc2);

        // Save the merged document.
        string outputPath = "MergedTables.docx";
        destDoc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        Console.WriteLine($"Merged document created successfully at: {Path.GetFullPath(outputPath)}");
    }
}
