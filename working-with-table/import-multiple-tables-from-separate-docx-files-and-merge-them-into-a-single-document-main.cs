using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the temporary source documents and the final merged document.
        string sourcePath1 = Path.Combine(outputDir, "Source1.docx");
        string sourcePath2 = Path.Combine(outputDir, "Source2.docx");
        string mergedPath = Path.Combine(outputDir, "Merged.docx");

        // -----------------------------------------------------------------
        // Create first source document with a simple 2x2 table.
        // -----------------------------------------------------------------
        Document sourceDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(sourceDoc1);
        builder1.Writeln("Source Document 1");
        Table table1 = builder1.StartTable();
        builder1.InsertCell();
        builder1.Write("Doc1 R1C1");
        builder1.InsertCell();
        builder1.Write("Doc1 R1C2");
        builder1.EndRow();
        builder1.InsertCell();
        builder1.Write("Doc1 R2C1");
        builder1.InsertCell();
        builder1.Write("Doc1 R2C2");
        builder1.EndRow();
        builder1.EndTable();
        sourceDoc1.Save(sourcePath1);

        // -----------------------------------------------------------------
        // Create second source document with a simple 3x1 table.
        // -----------------------------------------------------------------
        Document sourceDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(sourceDoc2);
        builder2.Writeln("Source Document 2");
        Table table2 = builder2.StartTable();
        builder2.InsertCell();
        builder2.Write("Doc2 Row1");
        builder2.EndRow();
        builder2.InsertCell();
        builder2.Write("Doc2 Row2");
        builder2.EndRow();
        builder2.InsertCell();
        builder2.Write("Doc2 Row3");
        builder2.EndRow();
        builder2.EndTable();
        sourceDoc2.Save(sourcePath2);

        // -----------------------------------------------------------------
        // Create the destination document that will hold all imported tables.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("Merged Document");
        destBuilder.Writeln(); // Add an empty paragraph to separate from imported tables.

        // -----------------------------------------------------------------
        // Helper method to import all tables from a source document.
        // -----------------------------------------------------------------
        void ImportTables(Document src)
        {
            // Find all Table nodes in the source document.
            NodeCollection tables = src.GetChildNodes(NodeType.Table, true);
            foreach (Table srcTable in tables)
            {
                // Import the table into the destination document preserving its formatting.
                NodeImporter importer = new NodeImporter(src, destDoc, ImportFormatMode.KeepSourceFormatting);
                Table importedTable = (Table)importer.ImportNode(srcTable, true);

                // Append the imported table to the body of the destination document.
                destDoc.FirstSection.Body.AppendChild(importedTable);
                // Add a paragraph after each table for visual separation.
                destDoc.FirstSection.Body.AppendChild(new Paragraph(destDoc));
            }
        }

        // Import tables from both source documents.
        ImportTables(sourceDoc1);
        ImportTables(sourceDoc2);

        // Save the merged document.
        destDoc.Save(mergedPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not saved correctly.");
    }
}
