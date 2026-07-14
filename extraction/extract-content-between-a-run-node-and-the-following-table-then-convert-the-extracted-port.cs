using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample source document containing a run of text followed by a table.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // First paragraph with some text.
        builder.Writeln("Paragraph before the target run.");

        // Second paragraph with the target run.
        builder.Write("TargetRunText "); // This creates a Run inside the current paragraph.

        // End the paragraph.
        builder.Writeln();

        // Insert a table after the run.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Save the source document to a local file.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -----------------------------------------------------------------
        // Load the document for extraction.
        // -----------------------------------------------------------------
        Document loaded = new Document(sourcePath);

        // Locate the run that contains the marker text "TargetRunText".
        Run runNode = null;
        foreach (Run run in loaded.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text != null && run.Text.Contains("TargetRunText"))
            {
                runNode = run;
                break;
            }
        }

        if (runNode == null)
            throw new InvalidOperationException("Target run node not found.");

        // Find the first table node that appears after the paragraph containing the run.
        Paragraph startParagraph = runNode.ParentParagraph;
        Node nextNode = startParagraph.NextSibling;
        Table tableNode = null;

        while (nextNode != null)
        {
            if (nextNode.NodeType == NodeType.Table)
            {
                tableNode = (Table)nextNode;
                break;
            }
            nextNode = nextNode.NextSibling;
        }

        if (tableNode == null)
            throw new InvalidOperationException("Following table not found.");

        // -----------------------------------------------------------------
        // Build a new document that will contain the extracted content.
        // Use NodeImporter to import nodes from the source document into the new document.
        // -----------------------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren();

        Section section = new Section(result);
        result.AppendChild(section);

        Body body = new Body(result);
        section.AppendChild(body);

        // Import the paragraph that contains the run (preserves styling).
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);
        Paragraph importedParagraph = (Paragraph)importer.ImportNode(startParagraph, true);
        body.AppendChild(importedParagraph);

        // Import the table.
        Table importedTable = (Table)importer.ImportNode(tableNode, true);
        body.AppendChild(importedTable);

        // Save the extracted portion as XPS.
        const string outputPath = "extracted.xps";
        result.Save(outputPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The XPS output file was not created.");
    }
}
