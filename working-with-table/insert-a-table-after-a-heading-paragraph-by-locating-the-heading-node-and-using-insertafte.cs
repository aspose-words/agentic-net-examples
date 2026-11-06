using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document that will contain the heading.
        Document mainDoc = new Document();
        DocumentBuilder mainBuilder = new DocumentBuilder(mainDoc);

        // Insert a heading paragraph (Heading 1 style).
        mainBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        mainBuilder.Writeln("Sample Heading");

        // Locate the heading paragraph node.
        Paragraph headingParagraph = (Paragraph)mainDoc.GetChild(NodeType.Paragraph, 0, true);
        if (headingParagraph == null || headingParagraph.ParagraphFormat.StyleIdentifier != StyleIdentifier.Heading1)
            throw new InvalidOperationException("Heading paragraph not found.");

        // Create a separate temporary document to build the table.
        Document tempDoc = new Document();
        DocumentBuilder tableBuilder = new DocumentBuilder(tempDoc);

        // Build a simple 2x2 table.
        tableBuilder.StartTable();

        tableBuilder.InsertCell();
        tableBuilder.Write("Cell 1");
        tableBuilder.EndRow();

        tableBuilder.InsertCell();
        tableBuilder.Write("Cell 2");
        tableBuilder.EndRow();

        tableBuilder.EndTable();

        // Retrieve the created table from the temporary document.
        Table sourceTable = (Table)tempDoc.GetChild(NodeType.Table, 0, true);
        if (sourceTable == null)
            throw new InvalidOperationException("Table was not created.");

        // Import the table into the main document.
        NodeImporter importer = new NodeImporter(tempDoc, mainDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(sourceTable, true);

        // Insert the imported table after the heading paragraph.
        headingParagraph.ParentNode.InsertAfter(importedTable, headingParagraph);

        // Save the resulting document.
        string outputPath = "Output.docx";
        mainDoc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
