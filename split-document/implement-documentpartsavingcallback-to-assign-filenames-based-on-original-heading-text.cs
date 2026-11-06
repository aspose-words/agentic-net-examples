using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = "SplitOutput";
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter One");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter one.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter Two");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter two.");

        // Save the whole document as a reference file
        string mainFilePath = Path.Combine(outputDir, "Document.html");
        HtmlSaveOptions mainSaveOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true
        };
        doc.Save(mainFilePath, mainSaveOptions);

        // Split the document by Heading1 paragraphs and save each part using the heading text as filename
        SplitByHeading(doc, outputDir);

        // Verify that files were created
        string[] generatedFiles = Directory.GetFiles(outputDir, "*.html");
        if (generatedFiles.Length == 0)
            throw new Exception("No HTML files were generated.");

        Console.WriteLine("Generated HTML files:");
        foreach (string file in generatedFiles)
            Console.WriteLine(file);
    }

    private static void SplitByHeading(Document sourceDoc, string outputDir)
    {
        // Collect all paragraphs in the document
        NodeCollection allParagraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);

        int i = 0;
        while (i < allParagraphs.Count)
        {
            Paragraph para = (Paragraph)allParagraphs[i];
            // Check if this paragraph is a Heading1
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                // Determine the range of nodes belonging to this heading
                int startIndex = i;
                int endIndex = startIndex + 1; // start after the heading
                while (endIndex < allParagraphs.Count)
                {
                    Paragraph nextPara = (Paragraph)allParagraphs[endIndex];
                    if (nextPara.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                        break; // next heading starts a new part
                    endIndex++;
                }

                // Create a new document for this part
                Document partDoc = new Document();
                // Ensure the destination document has a section to receive nodes
                Section destSection = partDoc.FirstSection;

                // Import and append nodes from startIndex to endIndex - 1
                for (int j = startIndex; j < endIndex; j++)
                {
                    Node importedNode = partDoc.ImportNode(allParagraphs[j], true, ImportFormatMode.KeepSourceFormatting);
                    destSection.Body.AppendChild(importedNode);
                }

                // Build a safe filename from the heading text
                string headingText = para.GetText().Trim();
                string safeName = MakeSafeFileName(headingText);
                string partFilePath = Path.Combine(outputDir, safeName + ".html");

                // Save the part as HTML
                HtmlSaveOptions partSaveOptions = new HtmlSaveOptions
                {
                    ExportImagesAsBase64 = true
                };
                partDoc.Save(partFilePath, partSaveOptions);

                // Move the index to the next heading (or end)
                i = endIndex;
            }
            else
            {
                i++;
            }
        }
    }

    private static string MakeSafeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name.Replace(' ', '_');
    }
}
