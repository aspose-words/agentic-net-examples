using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing heading paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Document Title");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter One");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter one.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter Two");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter two.");

        // Prepare the output folder.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Collect all heading paragraphs of level 1.
        List<Paragraph> headingParagraphs = new List<Paragraph>();
        NodeCollection allParagraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in allParagraphs)
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                headingParagraphs.Add(para);
        }

        if (headingParagraphs.Count == 0)
            throw new InvalidOperationException("No heading paragraphs were found to split the document.");

        // Split the document manually based on the collected headings.
        for (int i = 0; i < headingParagraphs.Count; i++)
        {
            Paragraph startHeading = headingParagraphs[i];
            Paragraph endHeading = (i + 1 < headingParagraphs.Count) ? headingParagraphs[i + 1] : null;

            // Create a new document that will hold the current part.
            Document partDoc = new Document();
            // Remove the default empty paragraph that Aspose.Words adds.
            partDoc.FirstSection.Body.RemoveAllChildren();

            // Import nodes from the source document into the part document.
            Node currentNode = startHeading;
            while (currentNode != null && currentNode != endHeading)
            {
                Node importedNode = partDoc.ImportNode(currentNode, true, ImportFormatMode.KeepSourceFormatting);
                partDoc.FirstSection.Body.AppendChild(importedNode);
                currentNode = currentNode.NextSibling;
            }

            // Build a safe file name from the heading text.
            string headingText = startHeading.GetText().Trim();
            foreach (char invalid in Path.GetInvalidFileNameChars())
                headingText = headingText.Replace(invalid, '_');

            if (string.IsNullOrWhiteSpace(headingText))
                headingText = $"Part{i + 1}";

            string fileName = $"{headingText}.html";
            string outPath = Path.Combine(outputDir, fileName);

            // Save the split part as HTML.
            partDoc.Save(outPath, SaveOptions.CreateSaveOptions(SaveFormat.Html));
        }

        // Verify that at least one split HTML file was generated.
        string[] htmlFiles = Directory.GetFiles(outputDir, "*.html");
        if (htmlFiles.Length == 0)
            throw new InvalidOperationException("No split HTML files were generated.");

        // Output the names of the generated files.
        foreach (string file in htmlFiles)
            Console.WriteLine($"Generated file: {Path.GetFileName(file)}");
    }
}
