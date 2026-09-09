using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample source document containing a run and a table.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph before the run (just for context).
        builder.Writeln("Paragraph before the run.");

        // The run that we will later extract.
        builder.Write("Extracted run text.");

        // End the paragraph that contains the run.
        builder.Writeln();

        // Insert a simple table after the run.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.Write("Cell B1");
        builder.EndRow();
        builder.EndTable();

        // Save the source document locally.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // ------------------------------------------------------------
        // 2. Load the document for extraction.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the first Run node (the one we created above).
        Run runNode = loadedDoc.GetChildNodes(NodeType.Run, true)[0] as Run;
        if (runNode == null)
            throw new InvalidOperationException("Run node not found.");

        // Locate the first Table node that follows the run.
        Table tableNode = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (tableNode == null)
            throw new InvalidOperationException("Table node not found.");

        // ------------------------------------------------------------
        // 3. Build a new document that will contain the extracted content.
        // ------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        // Create a new section and body for the result document.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // ------------------------------------------------------------
        // 4. Import the Run and Table nodes into the result document.
        // ------------------------------------------------------------
        // Use NodeImporter to copy nodes from the source document to the destination document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        // Import the run (inline node) and place it inside a new paragraph.
        Node importedRun = importer.ImportNode(runNode, true);
        Paragraph runParagraph = new Paragraph(resultDoc);
        runParagraph.AppendChild(importedRun);
        resultBody.AppendChild(runParagraph);

        // Import the table (block node) and append it directly to the body.
        Node importedTable = importer.ImportNode(tableNode, true);
        resultBody.AppendChild(importedTable);

        // ------------------------------------------------------------
        // 5. Save the extracted portion as XPS.
        // ------------------------------------------------------------
        const string outputPath = "extracted.xps";
        resultDoc.Save(outputPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The XPS output file was not created.");
    }
}
