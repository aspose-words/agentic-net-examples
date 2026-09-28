using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // Create a source document containing two tables with different styles.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First table – style "Table Grid".
        builder.Writeln("First Table:");
        builder.StartTable();
        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("A2");
        builder.InsertCell();
        builder.Write("B2");
        builder.EndRow();
        builder.EndTable();

        // Apply style to the first table.
        Table firstTable = (Table)sourceDoc.GetChildNodes(NodeType.Table, true)[0];
        firstTable.StyleIdentifier = StyleIdentifier.TableGrid;

        // Second table – style "Light List Accent 1".
        builder.Writeln();
        builder.Writeln("Second Table:");
        builder.StartTable();
        builder.InsertCell();
        builder.Write("C1");
        builder.InsertCell();
        builder.Write("D1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("C2");
        builder.InsertCell();
        builder.Write("D2");
        builder.EndRow();
        builder.EndTable();

        // Apply style to the second table.
        Table secondTable = (Table)sourceDoc.GetChildNodes(NodeType.Table, true)[1];
        secondTable.StyleIdentifier = StyleIdentifier.LightListAccent1;

        // Save the source document for reference.
        sourceDoc.Save("SourceDocument.docx");

        // ------------------------------------------------------------
        // Export each table to a separate DOCX file while preserving its style.
        // ------------------------------------------------------------
        NodeCollection tables = sourceDoc.GetChildNodes(NodeType.Table, true);
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = (Table)tables[i];

            // Create a new empty document that will hold the single table.
            Document destDoc = new Document();

            // Import the table from the source document into the destination document.
            NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);
            Table importedTable = (Table)importer.ImportNode(table, true);

            // Append the imported table to the body of the destination document.
            // A Body can contain a Table directly, so no need for a placeholder paragraph.
            destDoc.FirstSection.Body.AppendChild(importedTable);

            // Save the individual table document.
            string fileName = $"Table_{i + 1}.docx";
            destDoc.Save(fileName);

            // Verify that the file was created.
            if (!File.Exists(fileName))
                throw new Exception($"Failed to save {fileName}");
        }
    }
}
