using System;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // ---------- Create a source document with sample paragraphs ----------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("First paragraph in source document.");
        srcBuilder.Writeln("Second paragraph in source document.");
        srcBuilder.Writeln("Third paragraph in source document.");

        // ---------- Extract the paragraph nodes from the source document ----------
        ParagraphCollection sourceParagraphs = sourceDoc.FirstSection.Body.Paragraphs;

        // ---------- Create a new destination document ----------
        Document destDoc = new Document();
        // Remove the automatically created empty paragraph so the body starts empty.
        destDoc.FirstSection.Body.RemoveAllChildren();

        // ---------- Prepare a NodeImporter for efficient node copying ----------
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // ---------- Prepend the extracted paragraphs to the destination document ----------
        // Insert paragraphs in reverse order using PrependChild so that the original order is preserved.
        for (int i = sourceParagraphs.Count - 1; i >= 0; i--)
        {
            Node importedNode = importer.ImportNode(sourceParagraphs[i], true);
            // The body is a CompositeNode; PrependChild adds the node at the beginning.
            destDoc.FirstSection.Body.PrependChild(importedNode);
        }

        // ---------- Save the resulting document ----------
        destDoc.Save("PrependedDocument.docx", SaveFormat.Docx);
    }
}
