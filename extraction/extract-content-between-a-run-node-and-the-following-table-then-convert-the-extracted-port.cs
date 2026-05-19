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

        builder.Writeln("Intro paragraph.");

        // Paragraph that will contain the target run.
        builder.Writeln("This paragraph contains the target run:");
        Run targetRun = new Run(source, "TARGET_RUN");
        builder.CurrentParagraph.AppendChild(targetRun);
        builder.Writeln(); // End the paragraph.

        // Insert a table after the run.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the source document locally.
        const string sourcePath = "source.docx";
        source.Save(sourcePath);

        // -------------------------------------------------
        // 2. Load the source document and locate the required nodes.
        // -------------------------------------------------
        Document loaded = new Document(sourcePath);

        Run runNode = null;
        foreach (Run run in loaded.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "TARGET_RUN")
            {
                runNode = run;
                break;
            }
        }

        Table tableNode = null;
        // Find the first table that appears after the run in document order.
        if (runNode != null)
        {
            NodeCollection allNodes = loaded.GetChildNodes(NodeType.Any, true);
            int runIndex = allNodes.IndexOf(runNode);

            foreach (Table tbl in loaded.GetChildNodes(NodeType.Table, true))
            {
                int tblIndex = allNodes.IndexOf(tbl);
                if (tblIndex > runIndex)
                {
                    tableNode = tbl;
                    break;
                }
            }
        }

        if (runNode == null || tableNode == null)
            throw new InvalidOperationException("Required run or table node was not found.");

        // -------------------------------------------------
        // 3. Build a new document that will contain the extracted content.
        // -------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren();

        Section section = new Section(result);
        result.AppendChild(section);

        Body body = new Body(result);
        section.AppendChild(body);

        // -------------------------------------------------
        // 4. Import the run and table from the source document into the new document.
        //    Use NodeImporter to avoid cross‑document ownership errors.
        // -------------------------------------------------
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        // Import the run (inline node) and place it inside a new paragraph.
        Paragraph para = new Paragraph(result);
        body.AppendChild(para);
        Node importedRun = importer.ImportNode(runNode, true);
        para.AppendChild(importedRun);

        // Import the table (block node) and append it directly to the body.
        Node importedTable = importer.ImportNode(tableNode, true);
        body.AppendChild(importedTable);

        // -------------------------------------------------
        // 5. Save the extracted content as XPS.
        // -------------------------------------------------
        const string outputPath = "extracted.xps";
        result.Save(outputPath, SaveFormat.Xps);

        // Verify that the XPS file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The XPS output file was not created.");
    }
}
