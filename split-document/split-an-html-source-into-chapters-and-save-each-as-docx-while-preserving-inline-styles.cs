using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        if (Directory.Exists(outputDir))
            Directory.Delete(outputDir, true);
        Directory.CreateDirectory(outputDir);

        // Sample HTML content with headings (chapters) and inline styles.
        string htmlContent = @"
<!DOCTYPE html>
<html>
<head><title>Sample Document</title></head>
<body>
<h1>Chapter 1</h1>
<p>This is the <span style='color:red;'>first</span> paragraph of chapter 1.</p>
<p>Another paragraph with <b>bold</b> text.</p>
<h2>Section 1.1</h2>
<p>Content of a sub‑section.</p>
<h1>Chapter 2</h1>
<p>Paragraph in chapter 2 with <i>italic</i> style.</p>
<h1>Chapter 3</h1>
<p>Final chapter paragraph.</p>
</body>
</html>";

        // Load the HTML into an Aspose.Words Document.
        using (MemoryStream htmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(htmlContent)))
        {
            // LoadOptions can be used to specify the load format explicitly.
            LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Html };
            Document sourceDoc = new Document(htmlStream, loadOptions);

            // Collect all paragraphs in the document.
            NodeCollection paragraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);

            int chapterIndex = 0;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                Paragraph para = (Paragraph)paragraphs[i];

                // Identify chapter start: a paragraph styled as Heading 1.
                if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                {
                    chapterIndex++;

                    // Create a new empty document for the current chapter.
                    Document chapterDoc = new Document();
                    // Remove the default empty section created by the constructor.
                    chapterDoc.RemoveAllChildren();

                    // Add a new section with a body to hold imported nodes.
                    Section chapterSection = new Section(chapterDoc);
                    chapterDoc.AppendChild(chapterSection);
                    Body chapterBody = new Body(chapterDoc);
                    chapterSection.AppendChild(chapterBody);

                    // Prepare an importer to copy nodes while preserving source formatting.
                    NodeImporter importer = new NodeImporter(sourceDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);

                    // Import the heading paragraph itself.
                    Node importedNode = importer.ImportNode(para, true);
                    chapterBody.AppendChild(importedNode);

                    // Import subsequent nodes until the next Heading 1 or end of document.
                    int j = i + 1;
                    while (j < paragraphs.Count)
                    {
                        Paragraph nextPara = (Paragraph)paragraphs[j];
                        if (nextPara.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                            break; // Next chapter starts here.

                        // Import the node (could be a paragraph, table, etc.).
                        Node imported = importer.ImportNode(nextPara, true);
                        chapterBody.AppendChild(imported);
                        j++;
                    }

                    // Save the chapter as DOCX.
                    string chapterPath = Path.Combine(outputDir, $"Chapter_{chapterIndex}.docx");
                    chapterDoc.Save(chapterPath, SaveFormat.Docx);

                    // Advance the outer loop to the node before the next heading.
                    i = j - 1;
                }
            }

            // Simple validation: ensure at least one chapter file was created.
            string[] generatedFiles = Directory.GetFiles(outputDir, "*.docx");
            if (generatedFiles.Length == 0)
                throw new InvalidOperationException("No chapter files were generated.");

            // Optionally, output the list of generated files to the console.
            Console.WriteLine("Generated chapter files:");
            foreach (string file in generatedFiles)
                Console.WriteLine(file);
        }
    }
}
