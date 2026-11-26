using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractionBetweenHeadings
{
    public static void Main()
    {
        // Create a source document with headings and content between them.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Heading 1");
        srcBuilder.Writeln("Content before target range.");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        srcBuilder.Writeln("Heading 2");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        srcBuilder.Writeln("Paragraph A inside range.");
        srcBuilder.Writeln("Paragraph B inside range.");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        srcBuilder.Writeln("Heading 3");

        srcBuilder.Writeln("Content after target range.");
        sourceDoc.Save("source.docx");

        // Load the source document.
        Document loadedSource = new Document("source.docx");

        // Locate the start and end heading paragraphs.
        Paragraph startHeading = null;
        Paragraph endHeading = null;
        foreach (Paragraph para in loadedSource.FirstSection.Body.Paragraphs)
        {
            string text = para.GetText().Trim();
            if (text == "Heading 2")
                startHeading = para;
            else if (text == "Heading 3")
                endHeading = para;
        }

        if (startHeading == null || endHeading == null)
            throw new InvalidOperationException("Required headings were not found.");

        // Extract nodes between the two headings (excluding the headings themselves).
        List<Node> extractedNodes = new List<Node>();
        Node current = startHeading.NextSibling;
        while (current != null && current != endHeading)
        {
            extractedNodes.Add(current);
            current = current.NextSibling;
        }

        if (extractedNodes.Count == 0)
            throw new InvalidOperationException("No content found between the specified headings.");

        // Create a template document with a placeholder paragraph.
        Document templateDoc = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(templateDoc);
        tmplBuilder.Writeln("Template Header");
        tmplBuilder.Writeln("[Insert Extracted Content Here]");
        tmplBuilder.Writeln("Template Footer");
        templateDoc.Save("template.docx");

        // Load the template document.
        Document resultDoc = new Document("template.docx");

        // Find the placeholder paragraph.
        Paragraph placeholder = null;
        foreach (Paragraph para in resultDoc.FirstSection.Body.Paragraphs)
        {
            if (para.GetText().Contains("[Insert Extracted Content Here]"))
            {
                placeholder = para;
                break;
            }
        }

        if (placeholder == null)
            throw new InvalidOperationException("Placeholder paragraph not found in the template.");

        // Insert extracted content at the placeholder position.
        Body body = resultDoc.FirstSection.Body;
        Node insertAfterNode = placeholder.PreviousSibling;
        // Remove the placeholder.
        body.RemoveChild(placeholder);

        foreach (Node node in extractedNodes)
        {
            // Import the node into the result document and clone it.
            Node importedNode = resultDoc.ImportNode(node, true);
            if (insertAfterNode == null)
            {
                // Insert at the beginning of the body.
                body.PrependChild(importedNode);
                insertAfterNode = importedNode;
            }
            else
            {
                body.InsertAfter(importedNode, insertAfterNode);
                insertAfterNode = importedNode;
            }
        }

        // Save the final document.
        resultDoc.Save("result.docx");

        // Validate that the result file was created.
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("Result document was not created.");

        // Optional: indicate success (no interactive output required).
    }
}
