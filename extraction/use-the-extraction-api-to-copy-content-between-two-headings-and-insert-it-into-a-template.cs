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
        // 1. Create a source document with several headings and content.
        // -----------------------------------------------------------------
        Document source = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(source);

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        srcBuilder.Writeln("Heading 1");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        srcBuilder.Writeln("Heading 2");
        srcBuilder.Writeln("Paragraph under Heading 2.");

        srcBuilder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        srcBuilder.Writeln("Heading 3");
        srcBuilder.Writeln("Paragraph under Heading 3.");

        source.Save("source.docx");

        // -----------------------------------------------------------------
        // 2. Create a template document that will receive the extracted content.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);

        tmplBuilder.Writeln("Template start");
        tmplBuilder.Writeln("[Placeholder]"); // marker where content will be inserted
        tmplBuilder.Writeln("Template end");

        template.Save("template.docx");

        // -----------------------------------------------------------------
        // 3. Load both documents.
        // -----------------------------------------------------------------
        Document srcDoc = new Document("source.docx");
        Document tmplDoc = new Document("template.docx");

        // -----------------------------------------------------------------
        // 4. Locate the start and end heading paragraphs.
        // -----------------------------------------------------------------
        Paragraph startHeading = null;
        Paragraph endHeading = null;

        foreach (Paragraph para in srcDoc.FirstSection.Body.Paragraphs)
        {
            string text = para.GetText().Trim();
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 && text == "Heading 1")
                startHeading = para;
            else if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading3 && text == "Heading 3")
                endHeading = para;
        }

        if (startHeading == null || endHeading == null)
            throw new InvalidOperationException("Required headings were not found in the source document.");

        // -----------------------------------------------------------------
        // 5. Collect all nodes that lie between the two headings (exclusive).
        // -----------------------------------------------------------------
        List<Node> nodesBetween = new List<Node>();
        Node curNode = startHeading.NextSibling;

        while (curNode != null && curNode != endHeading)
        {
            Node next = curNode.NextSibling; // preserve next reference before moving
            nodesBetween.Add(curNode);
            curNode = next;
        }

        if (nodesBetween.Count == 0)
            throw new InvalidOperationException("No content found between the specified headings.");

        // -----------------------------------------------------------------
        // 6. Find the placeholder paragraph in the template.
        // -----------------------------------------------------------------
        Paragraph placeholder = null;
        foreach (Paragraph para in tmplDoc.FirstSection.Body.Paragraphs)
        {
            if (para.GetText().Contains("[Placeholder]"))
            {
                placeholder = para;
                break;
            }
        }

        if (placeholder == null)
            throw new InvalidOperationException("Placeholder paragraph not found in the template document.");

        // -----------------------------------------------------------------
        // 7. Import the extracted nodes into the template after the placeholder.
        // -----------------------------------------------------------------
        NodeImporter importer = new NodeImporter(srcDoc, tmplDoc, ImportFormatMode.KeepSourceFormatting);
        CompositeNode destinationStory = placeholder.ParentNode as CompositeNode;
        Node insertionPoint = placeholder;

        foreach (Node node in nodesBetween)
        {
            Node importedNode = importer.ImportNode(node, true);
            destinationStory.InsertAfter(importedNode, insertionPoint);
            insertionPoint = importedNode; // advance insertion point
        }

        // Remove the placeholder paragraph itself.
        placeholder.Remove();

        // -----------------------------------------------------------------
        // 8. Save the resulting document.
        // -----------------------------------------------------------------
        tmplDoc.Save("result.docx");

        // -----------------------------------------------------------------
        // 9. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("Result document was not created.");
    }
}
