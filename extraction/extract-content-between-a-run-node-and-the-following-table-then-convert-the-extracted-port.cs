using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample source document with a target run and a table.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        builder.Writeln("Intro paragraph.");
        builder.Writeln("Paragraph before target run.");

        // Write some runs; the target run contains unique text.
        builder.Write("Some ");
        builder.Font.Bold = true;
        builder.Write("TargetRun");
        builder.Font.Bold = false;
        builder.Writeln(" after run.");

        // Insert a table after the paragraph containing the target run.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell1");
        builder.InsertCell();
        builder.Write("Cell2");
        builder.EndRow();
        builder.EndTable();

        // Save the source document to a local file.
        const string sourcePath = "input.docx";
        source.Save(sourcePath);

        // ---------------------------------------------------------------
        // 2. Load the document for processing and locate the target run.
        // ---------------------------------------------------------------
        Document doc = new Document(sourcePath);

        Run targetRun = null;
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "TargetRun")
            {
                targetRun = run;
                break;
            }
        }

        if (targetRun == null)
            throw new InvalidOperationException("Target run not found.");

        // ---------------------------------------------------------------
        // 3. Find the first table that follows the target run.
        // ---------------------------------------------------------------
        Node node = targetRun;
        Table followingTable = null;
        while (node != null && followingTable == null)
        {
            // NextPreOrder requires the root node of the document tree.
            node = node.NextPreOrder(doc);
            if (node != null && node.NodeType == NodeType.Table)
                followingTable = (Table)node;
        }

        if (followingTable == null)
            throw new InvalidOperationException("Following table not found.");

        // ---------------------------------------------------------------
        // 4. Build a new document that will contain the extracted content.
        // ---------------------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren();

        Section section = new Section(result);
        result.AppendChild(section);

        Body body = new Body(result);
        section.AppendChild(body);

        // The extracted run is an inline node; place it inside a new paragraph.
        Paragraph extractedParagraph = new Paragraph(result);
        // Import the run into the destination document to preserve styles.
        Run importedRun = (Run)result.ImportNode(targetRun, true);
        extractedParagraph.AppendChild(importedRun);
        body.AppendChild(extractedParagraph);

        // Import the table into the destination document.
        Table importedTable = (Table)result.ImportNode(followingTable, true);
        body.AppendChild(importedTable);

        // ---------------------------------------------------------------
        // 5. Save the extracted portion as XPS.
        // ---------------------------------------------------------------
        const string outputPath = "extracted.xps";
        result.Save(outputPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The XPS output file was not created.");
    }
}
