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
        // 1. Create a source document that contains two headings and some
        //    content between them.
        // -----------------------------------------------------------------
        string sourcePath = "source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First heading (start marker)
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Start Heading");

        // Content to be extracted
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph 1 between headings.");
        srcBuilder.Writeln("Paragraph 2 between headings.");

        // Insert a simple table as part of the extracted range
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

        // Second heading (end marker)
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("End Heading");

        // Save the source document
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document and locate the two heading paragraphs.
        // -----------------------------------------------------------------
        Document loadedSource = new Document(sourcePath);
        Paragraph startHeading = null;
        Paragraph endHeading = null;

        foreach (Paragraph para in loadedSource.GetChildNodes(NodeType.Paragraph, true))
        {
            string text = para.GetText().Trim();
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                if (text == "Start Heading")
                    startHeading = para;
                else if (text == "End Heading")
                    endHeading = para;
            }
        }

        if (startHeading == null || endHeading == null)
            throw new InvalidOperationException("Required headings were not found in the source document.");

        // -----------------------------------------------------------------
        // 3. Collect all nodes that lie between the two headings (exclusive).
        // -----------------------------------------------------------------
        List<Node> nodesToExtract = new List<Node>();
        Node current = startHeading.NextSibling;
        while (current != null && current != endHeading)
        {
            Node next = current.NextSibling; // Preserve next node before we move it
            nodesToExtract.Add(current);
            current = next;
        }

        if (nodesToExtract.Count == 0)
            throw new InvalidOperationException("No content found between the specified headings.");

        // -----------------------------------------------------------------
        // 4. Create a template document where the extracted content will be inserted.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("=== Template Header ===");
        tmplBuilder.Writeln("Content will be inserted below:");
        templateDoc.Save(templatePath);

        // Load the template
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 5. Import the extracted nodes into the template document.
        // -----------------------------------------------------------------
        // We'll insert after the last paragraph of the template body.
        Paragraph insertionPoint = loadedTemplate.FirstSection.Body.LastParagraph;

        NodeImporter importer = new NodeImporter(loadedSource, loadedTemplate, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in nodesToExtract)
        {
            // Import the node (deep clone) into the destination document.
            Node importedNode = importer.ImportNode(node, true);
            // Insert after the current insertion point.
            insertionPoint.ParentNode.InsertAfter(importedNode, insertionPoint);
            // Update the insertion point so subsequent nodes are appended in order.
            insertionPoint = importedNode as Paragraph ?? insertionPoint;
        }

        // -----------------------------------------------------------------
        // 6. Save the resulting document.
        // -----------------------------------------------------------------
        string resultPath = "result.docx";
        loadedTemplate.Save(resultPath);

        // Validate that the result file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");

        // Optional: clean up intermediate files (comment out if you want to inspect them)
        // File.Delete(sourcePath);
        // File.Delete(templatePath);
    }
}
