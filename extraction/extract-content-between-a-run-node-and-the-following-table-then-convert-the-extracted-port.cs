using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with a run and a table.
        // -------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        // First paragraph with a run.
        builder.Writeln("Intro paragraph.");
        builder.Write("Run text "); // This creates a Run inside the current paragraph.
        builder.Writeln(); // End the paragraph.

        // Insert a simple table after the run.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Save the source document locally.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the document for extraction.
        // -------------------------------------------------
        Document loaded = new Document(sourcePath);

        // Locate the first Run node.
        Paragraph firstParagraph = loaded.FirstSection.Body.Paragraphs[0];
        Run runNode = firstParagraph.Runs.Count > 0 ? firstParagraph.Runs[0] : null;
        if (runNode == null)
            throw new InvalidOperationException("Run node not found.");

        // Locate the first Table node that follows the run.
        Table tableNode = loaded.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (tableNode == null)
            throw new InvalidOperationException("Table node not found.");

        // -------------------------------------------------
        // 3. Build a new document that will contain the extracted content.
        // -------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren(); // Remove the default empty section/paragraph.

        // Create a new section and body.
        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        // -------------------------------------------------
        // 4. Import the Run and Table nodes into the new document.
        // -------------------------------------------------
        // NodeImporter handles style and list translation between documents.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        // Import the Run node.
        Node importedRun = importer.ImportNode(runNode, true);
        // Wrap the imported Run inside a new Paragraph (inline nodes must be inside a Paragraph).
        Paragraph extractedParagraph = new Paragraph(result);
        extractedParagraph.AppendChild(importedRun);
        body.AppendChild(extractedParagraph);

        // Import the Table node.
        Node importedTable = importer.ImportNode(tableNode, true);
        body.AppendChild(importedTable);

        // -------------------------------------------------
        // 5. Save the extracted portion as XPS.
        // -------------------------------------------------
        const string outputPath = "extracted.xps";
        result.Save(outputPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the XPS output file.");
    }
}
