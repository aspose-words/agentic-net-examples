using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document with sample content.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("First paragraph.");
        srcBuilder.Writeln("Second paragraph.");
        srcBuilder.Writeln("Third paragraph.");
        sourceDoc.Save("source.docx");

        // Load the source document.
        Document loadedSource = new Document("source.docx");

        // Extract all paragraph nodes from the source document.
        NodeCollection sourceParagraphs = loadedSource.GetChildNodes(NodeType.Paragraph, true);

        // Prepare a new destination document.
        Document destDoc = new Document();
        // Remove the default empty section/paragraph.
        destDoc.RemoveAllChildren();

        // Create a new section and body for the destination document.
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Importer to handle style and list translation.
        NodeImporter importer = new NodeImporter(loadedSource, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Prepend the extracted paragraphs to the beginning of the destination body.
        // Iterate in reverse order to preserve the original sequence when using PrependChild.
        for (int i = sourceParagraphs.Count - 1; i >= 0; i--)
        {
            Node srcNode = sourceParagraphs[i];
            // Skip the final empty paragraph that Aspose.Words adds automatically.
            if (srcNode is Paragraph para && para.IsEndOfSection && !para.HasChildNodes)
                continue;

            Node importedNode = importer.ImportNode(srcNode, true);
            destBody.PrependChild(importedNode);
        }

        // Save the resulting document.
        destDoc.Save("result.docx");

        // Validate that the output file was created.
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("The result document was not created.");

        // Optional: indicate success (no interactive output required).
        Console.WriteLine("Extraction and prepend operation completed successfully.");
    }
}
