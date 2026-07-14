using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Intro paragraph.");
        builder.Writeln("Start extraction paragraph.");

        // Insert a simple table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("End of document.");
        string inputPath = "input.docx";
        source.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document for extraction.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // Locate the start paragraph (second paragraph) and the target table.
        Paragraph startParagraph = loaded.FirstSection.Body.Paragraphs[1];
        Table targetTable = loaded.GetChildNodes(NodeType.Table, true)[0] as Table;

        if (startParagraph == null || targetTable == null)
            throw new InvalidOperationException("Required extraction nodes were not found.");

        // -----------------------------------------------------------------
        // 3. Verify that the start node appears before the end node.
        // -----------------------------------------------------------------
        NodeCollection bodyChildren = loaded.FirstSection.Body.GetChildNodes(NodeType.Any, false);
        int startIndex = bodyChildren.IndexOf(startParagraph);
        int endIndex = bodyChildren.IndexOf(targetTable);

        if (startIndex == -1 || endIndex == -1)
            throw new InvalidOperationException("Unable to determine node positions.");

        if (startIndex > endIndex)
            throw new InvalidOperationException("Start node appears after the end node; extraction aborted.");

        // -----------------------------------------------------------------
        // 4. Build a new document containing the extracted nodes.
        // -----------------------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren();

        Section resultSection = new Section(result);
        result.AppendChild(resultSection);

        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // Import nodes from the source document into the destination document.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        Node importedParagraph = importer.ImportNode(startParagraph, true);
        Node importedTable = importer.ImportNode(targetTable, true);

        resultBody.AppendChild(importedParagraph);
        resultBody.AppendChild(importedTable);

        // -----------------------------------------------------------------
        // 5. Save the extracted content and validate the output.
        // -----------------------------------------------------------------
        string outputPath = "extracted.docx";
        result.Save(outputPath);

        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Extraction output file was not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
