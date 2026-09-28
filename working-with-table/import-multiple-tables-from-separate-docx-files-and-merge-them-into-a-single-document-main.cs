using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Paths for temporary source and output documents.
        string srcPath1 = Path.Combine(Directory.GetCurrentDirectory(), "Source1.docx");
        string srcPath2 = Path.Combine(Directory.GetCurrentDirectory(), "Source2.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Merged.docx");

        // Create first source document with a simple table.
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("Source Document 1");
        builder1.StartTable();
        // Row 1
        builder1.InsertCell();
        builder1.Write("Table1 - Row1, Cell1");
        builder1.InsertCell();
        builder1.Write("Table1 - Row1, Cell2");
        builder1.EndRow();
        // Row 2
        builder1.InsertCell();
        builder1.Write("Table1 - Row2, Cell1");
        builder1.InsertCell();
        builder1.Write("Table1 - Row2, Cell2");
        builder1.EndRow();
        builder1.EndTable();
        srcDoc1.Save(srcPath1);

        // Create second source document with a different table.
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Source Document 2");
        builder2.StartTable();
        // Row 1
        builder2.InsertCell();
        builder2.Write("Table2 - Row1, Cell1");
        builder2.InsertCell();
        builder2.Write("Table2 - Row1, Cell2");
        builder2.EndRow();
        // Row 2
        builder2.InsertCell();
        builder2.Write("Table2 - Row2, Cell1");
        builder2.InsertCell();
        builder2.Write("Table2 - Row2, Cell2");
        builder2.EndRow();
        builder2.EndTable();
        srcDoc2.Save(srcPath2);

        // Create the destination document where tables will be merged.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Merged Document");
        destBuilder.Writeln(); // Add an empty paragraph before tables.

        // Helper method to import all tables from a source document.
        void ImportTables(Document source)
        {
            // Prepare an importer that keeps source formatting.
            NodeImporter importer = new NodeImporter(source, destDoc, ImportFormatMode.KeepSourceFormatting);
            // Find all tables in the source document.
            NodeCollection tables = source.GetChildNodes(NodeType.Table, true);
            foreach (Table table in tables)
            {
                // Import the table into the destination document.
                Node importedNode = importer.ImportNode(table, true);
                // Append the imported table to the body of the destination document.
                destDoc.FirstSection.Body.AppendChild(importedNode);
                // Add a paragraph separator after each imported table.
                destDoc.FirstSection.Body.AppendChild(new Paragraph(destDoc));
            }
        }

        // Load source documents (they are already in memory, but loading from file demonstrates the workflow).
        Document loadedSrc1 = new Document(srcPath1);
        Document loadedSrc2 = new Document(srcPath2);

        // Import tables from both source documents.
        ImportTables(loadedSrc1);
        ImportTables(loadedSrc2);

        // Save the merged document.
        destDoc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The merged document was not saved correctly.");
        }
    }
}
