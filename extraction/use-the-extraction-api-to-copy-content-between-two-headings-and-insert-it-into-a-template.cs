using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ExtractionExample
{
    public static void Main()
    {
        // Create a source document with two headings and some content between them.
        var sourceDoc = new Document();
        var srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Start Heading");
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph under start heading.");
        srcBuilder.Writeln("Another paragraph under start heading.");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("End Heading");
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph after end heading.");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Create a template document that contains a placeholder where the extracted content will be inserted.
        var templateDoc = new Document();
        var tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("Template Header");
        tmplBuilder.Writeln("[InsertHere]"); // Placeholder paragraph
        tmplBuilder.Writeln("Template Footer");

        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the source document for extraction.
        var loadedSource = new Document(sourcePath);

        // Locate the start and end heading paragraphs.
        Paragraph startHeading = null;
        Paragraph endHeading = null;
        foreach (Paragraph para in loadedSource.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                string text = para.GetText().Trim();
                if (text == "Start Heading")
                    startHeading = para;
                else if (text == "End Heading")
                    endHeading = para;
            }
        }

        if (startHeading == null || endHeading == null)
            throw new InvalidOperationException("Required headings were not found in the source document.");

        // Collect nodes that lie between the two headings (excluding the headings themselves).
        var nodesToCopy = new List<Node>();
        Node current = startHeading.NextSibling;
        while (current != null && current != endHeading)
        {
            nodesToCopy.Add(current);
            current = current.NextSibling;
        }

        if (nodesToCopy.Count == 0)
            throw new InvalidOperationException("No content found between the specified headings.");

        // Load the template document where the extracted content will be inserted.
        var loadedTemplate = new Document(templatePath);

        // Find the placeholder paragraph.
        Paragraph placeholder = null;
        foreach (Paragraph para in loadedTemplate.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.GetText().Contains("[InsertHere]"))
            {
                placeholder = para;
                break;
            }
        }

        if (placeholder == null)
            throw new InvalidOperationException("Placeholder paragraph not found in the template document.");

        // Prepare an importer to move nodes from the source document to the template document.
        var importer = new NodeImporter(loadedSource, loadedTemplate, ImportFormatMode.KeepSourceFormatting);

        // Insert imported nodes after the placeholder.
        Node insertAfterNode = placeholder;
        foreach (Node node in nodesToCopy)
        {
            Node importedNode = importer.ImportNode(node, true);
            insertAfterNode.ParentNode.InsertAfter(importedNode, insertAfterNode);
            insertAfterNode = importedNode;
        }

        // Remove the placeholder paragraph.
        placeholder.Remove();

        // Save the resulting document.
        const string resultPath = "result.docx";
        loadedTemplate.Save(resultPath);

        // Create a simple JSON report about the extraction.
        var report = new
        {
            ExtractedNodeCount = nodesToCopy.Count,
            SourceDocument = Path.GetFullPath(sourcePath),
            TemplateDocument = Path.GetFullPath(templatePath),
            ResultDocument = Path.GetFullPath(resultPath)
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        const string reportPath = "extraction-report.json";
        File.WriteAllText(reportPath, jsonReport);

        // Validate that output files were created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("Result document was not created.");
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Extraction report was not created.");

        Console.WriteLine("Extraction and insertion completed successfully.");
    }
}
