using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a source document with sample paragraphs.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("Paragraph 1");
        sourceBuilder.Writeln("Paragraph 2");
        sourceBuilder.Writeln("Paragraph 3");
        sourceBuilder.Writeln("Paragraph 4");
        sourceDoc.Save("source.docx");

        // -----------------------------------------------------------------
        // 2. Extract the first three paragraphs from the source document.
        // -----------------------------------------------------------------
        NodeCollection allParagraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);
        List<Node> extractedParagraphs = new List<Node>();
        int paragraphsToExtract = Math.Min(3, allParagraphs.Count);
        for (int i = 0; i < paragraphsToExtract; i++)
        {
            extractedParagraphs.Add(allParagraphs[i]);
        }

        // -----------------------------------------------------------------
        // 3. Create a new destination document and build a clean structure.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        // Remove the default section/paragraph that Aspose.Words creates.
        destDoc.RemoveAllChildren();

        // Add a new section with a body.
        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);
        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // Use NodeImporter to import nodes from the source document into the destination.
        NodeImporter importer = new NodeImporter(sourceDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Append the extracted paragraphs in their original order.
        foreach (Node paragraph in extractedParagraphs)
        {
            Node importedNode = importer.ImportNode(paragraph, true);
            destBody.AppendChild(importedNode);
        }

        // -----------------------------------------------------------------
        // 4. Save the destination document.
        // -----------------------------------------------------------------
        destDoc.Save("result.docx");

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists("result.docx"))
            throw new InvalidOperationException("The result document was not created.");

        Console.WriteLine("Document created successfully: result.docx");
    }
}
