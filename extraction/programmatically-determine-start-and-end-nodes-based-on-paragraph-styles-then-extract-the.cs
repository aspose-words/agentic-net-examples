using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ---------- Create a sample source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Create custom paragraph styles that will act as markers.
        Style startStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "StartStyle");
        startStyle.Font.Bold = true;

        Style endStyle = sourceDoc.Styles.Add(StyleType.Paragraph, "EndStyle");
        endStyle.Font.Italic = true;

        // Normal paragraph before the marked range.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Paragraph before start.");

        // Paragraph with the start style.
        builder.ParagraphFormat.StyleName = "StartStyle";
        builder.Writeln("This is the start paragraph.");

        // Some middle paragraphs.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Middle paragraph 1.");
        builder.Writeln("Middle paragraph 2.");

        // Paragraph with the end style.
        builder.ParagraphFormat.StyleName = "EndStyle";
        builder.Writeln("This is the end paragraph.");

        // Normal paragraph after the marked range.
        builder.ParagraphFormat.StyleName = "Normal";
        builder.Writeln("Paragraph after end.");

        // Save the source document locally.
        const string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // ---------- Load the document and locate start/end paragraphs ----------
        Document doc = new Document(sourcePath);

        Paragraph startParagraph = null;
        Paragraph endParagraph = null;

        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (int i = 0; i < allParagraphs.Count; i++)
        {
            Paragraph para = (Paragraph)allParagraphs[i];
            if (para.ParagraphFormat.StyleName == "StartStyle")
            {
                startParagraph = para;

                // Search forward for the end style.
                for (int j = i + 1; j < allParagraphs.Count; j++)
                {
                    Paragraph possibleEnd = (Paragraph)allParagraphs[j];
                    if (possibleEnd.ParagraphFormat.StyleName == "EndStyle")
                    {
                        endParagraph = possibleEnd;
                        break;
                    }
                }
                break;
            }
        }

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Could not locate start or end paragraph based on styles.");

        // ---------- Extract the range preserving original styling ----------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        // Build a minimal document structure: Section -> Body.
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Use NodeImporter to copy nodes while keeping source formatting.
        NodeImporter importer = new NodeImporter(doc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Walk from startParagraph to endParagraph inclusive.
        Node currentNode = startParagraph;
        while (true)
        {
            // Import the current node (which is a Paragraph) and add it to the destination body.
            Node importedNode = importer.ImportNode(currentNode, true);
            body.AppendChild(importedNode);

            if (currentNode == endParagraph)
                break;

            // Move to the next sibling that is a Paragraph.
            currentNode = currentNode.NextSibling;
            while (currentNode != null && currentNode.NodeType != NodeType.Paragraph)
                currentNode = currentNode.NextSibling;

            if (currentNode == null)
                break; // Safety check; should not happen if endParagraph is after startParagraph.
        }

        // Save the extracted segment.
        const string extractedPath = "Extracted.docx";
        extractedDoc.Save(extractedPath);

        Console.WriteLine($"Extraction completed. Extracted document saved as '{extractedPath}'.");
    }
}
