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

        // First heading
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Start Heading");

        // Content between the headings
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph 1 between headings.");
        srcBuilder.Writeln("Paragraph 2 between headings.");

        // Second heading
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("End Heading");

        // Additional content after the second heading (should not be extracted)
        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph after end heading.");

        source.Save("source.docx");

        // -----------------------------------------------------------------
        // 2. Load the source document and extract nodes between the two headings.
        // -----------------------------------------------------------------
        Document srcDoc = new Document("source.docx");
        Paragraph startHeading = null;
        Paragraph endHeading = null;

        // Locate the two heading paragraphs.
        foreach (Paragraph para in srcDoc.FirstSection.Body.Paragraphs)
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                if (startHeading == null)
                    startHeading = para;
                else
                {
                    endHeading = para;
                    break;
                }
            }
        }

        if (startHeading == null || endHeading == null)
            throw new InvalidOperationException("Both start and end headings must exist.");

        // Collect all nodes that lie between the two headings (exclusive).
        List<Node> extractedNodes = new List<Node>();
        Node curNode = startHeading.NextSibling;
        while (curNode != null && curNode != endHeading)
        {
            Node next = curNode.NextSibling; // Preserve next reference before any modifications.
            extractedNodes.Add(curNode);
            curNode = next;
        }

        if (extractedNodes.Count == 0)
            throw new InvalidOperationException("No content found between the specified headings.");

        // -----------------------------------------------------------------
        // 3. Create a template document that contains a placeholder paragraph.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);
        tmplBuilder.Writeln("Template Header");
        tmplBuilder.Writeln("[Insert Extracted Content Here]");
        tmplBuilder.Writeln("Template Footer");
        template.Save("template.docx");

        // -----------------------------------------------------------------
        // 4. Load the template and replace the placeholder with extracted content.
        // -----------------------------------------------------------------
        Document tmplDoc = new Document("template.docx");
        Body tmplBody = tmplDoc.FirstSection.Body;

        // Find the placeholder paragraph.
        Paragraph placeholder = null;
        foreach (Paragraph para in tmplBody.Paragraphs)
        {
            if (para.GetText().Contains("[Insert Extracted Content Here]"))
            {
                placeholder = para;
                break;
            }
        }

        if (placeholder == null)
            throw new InvalidOperationException("Placeholder paragraph not found in the template.");

        // Remove the placeholder paragraph.
        placeholder.Remove();

        // Import and insert each extracted node into the template body.
        NodeImporter importer = new NodeImporter(srcDoc, tmplDoc, ImportFormatMode.KeepSourceFormatting);
        foreach (Node node in extractedNodes)
        {
            Node importedNode = importer.ImportNode(node, true);
            tmplBody.AppendChild(importedNode);
        }

        // -----------------------------------------------------------------
        // 5. Save the resulting document.
        // -----------------------------------------------------------------
        tmplDoc.Save("result.docx");

        // Verify that the output file was created.
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("Result document was not created.");

        // Optional: output a simple confirmation to the console.
        Console.WriteLine("Extraction and insertion completed successfully.");
    }
}
