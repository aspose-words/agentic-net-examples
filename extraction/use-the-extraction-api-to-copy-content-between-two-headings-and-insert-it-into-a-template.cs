using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a source document with two headings and some content.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);

        // First heading – start marker.
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Start");

        // Content between the headings (paragraphs and a table).
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph 1 between headings.");
        srcBuilder.Writeln("Paragraph 2 between headings.");

        srcBuilder.StartTable();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell A1");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell B1");
        srcBuilder.EndRow();
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell A2");
        srcBuilder.InsertCell();
        srcBuilder.Write("Cell B2");
        srcBuilder.EndTable();

        // Second heading – end marker.
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("End");

        // Save the source document.
        source.Save("source.docx");

        // -----------------------------------------------------------------
        // 2. Load the source document and extract block-level nodes between the headings.
        // -----------------------------------------------------------------
        Document srcLoaded = new Document("source.docx");
        Body srcBody = srcLoaded.FirstSection.Body;

        // Locate the start and end heading paragraphs.
        Paragraph startHeading = null;
        Paragraph endHeading = null;
        foreach (Paragraph para in srcBody.Paragraphs)
        {
            string text = para.GetText().Trim();
            if (text == "Start")
                startHeading = para;
            else if (text == "End")
                endHeading = para;
        }

        if (startHeading == null || endHeading == null)
            throw new InvalidOperationException("Start or End heading not found.");

        // Determine the indexes of the headings within the body’s block-level collection.
        int startIdx = srcBody.Paragraphs.IndexOf(startHeading);
        int endIdx = srcBody.Paragraphs.IndexOf(endHeading);
        if (startIdx < 0 || endIdx < 0 || startIdx >= endIdx)
            throw new InvalidOperationException("Invalid heading positions.");

        // Collect block-level nodes (Paragraphs and Tables) that lie between the two headings.
        List<Node> extractedNodes = new List<Node>();

        // Paragraphs between the headings.
        for (int i = startIdx + 1; i < endIdx; i++)
        {
            Paragraph para = srcBody.Paragraphs[i];
            extractedNodes.Add(para);
        }

        // Tables that are positioned between the headings.
        foreach (Table table in srcBody.Tables)
        {
            // A table’s position can be determined by its index in the body’s child node collection.
            int tableIdx = srcBody.IndexOf(table);
            if (tableIdx > srcBody.IndexOf(startHeading) && tableIdx < srcBody.IndexOf(endHeading))
                extractedNodes.Add(table);
        }

        // -----------------------------------------------------------------
        // 3. Create a template document where the extracted content will be inserted.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);
        tmplBuilder.Writeln("=== Template Header ===");
        tmplBuilder.Writeln("Content will be inserted below:");
        // Save the template.
        template.Save("template.docx");

        // -----------------------------------------------------------------
        // 4. Load the template and import the extracted nodes.
        // -----------------------------------------------------------------
        Document tmplLoaded = new Document("template.docx");
        Body tmplBody = tmplLoaded.FirstSection.Body;

        // Use NodeImporter to import nodes from source to template.
        NodeImporter importer = new NodeImporter(srcLoaded, tmplLoaded, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in extractedNodes)
        {
            // Import the node (deep clone) into the destination document.
            Node imported = importer.ImportNode(node, true);
            // Append only block-level nodes (Paragraph or Table) to the body.
            if (imported.NodeType == NodeType.Paragraph || imported.NodeType == NodeType.Table)
                tmplBody.AppendChild(imported);
        }

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        tmplLoaded.Save("result.docx");

        // Validate that the result file was created.
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("Result document was not created.");

        Console.WriteLine("Extraction and insertion completed successfully.");
    }
}
