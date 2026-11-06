using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentByHeading
{
    public static void Main()
    {
        // Create a sample document with two Heading1 paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First heading and its content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");

        // Second heading and its content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Ensure the output directory exists.
        string outputDir = "SplitOutput";
        Directory.CreateDirectory(outputDir);

        // Split the document manually by Heading1 paragraphs.
        // Collect all paragraphs in the source document.
        NodeCollection allParagraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);

        int chapterIndex = 0; // Used for naming output files.

        for (int i = 0; i < allParagraphs.Count; i++)
        {
            Paragraph para = (Paragraph)allParagraphs[i];

            // Identify a Heading1 paragraph – this marks the start of a new chapter.
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
            {
                chapterIndex++;

                // Create a new empty document that will hold the current chapter.
                Document chapterDoc = new Document();
                Section chapterSection = chapterDoc.FirstSection; // The document already contains one empty section.
                Body chapterBody = chapterSection.Body;

                // Importer to copy nodes from the source document to the chapter document.
                NodeImporter importer = new NodeImporter(sourceDoc, chapterDoc, ImportFormatMode.KeepSourceFormatting);

                // Import the heading paragraph.
                chapterBody.AppendChild(importer.ImportNode(para, true));

                // Import following paragraphs until the next Heading1 or the end of the document.
                int j = i + 1;
                while (j < allParagraphs.Count)
                {
                    Paragraph nextPara = (Paragraph)allParagraphs[j];
                    if (nextPara.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                        break; // Next chapter starts here.

                    chapterBody.AppendChild(importer.ImportNode(nextPara, true));
                    j++;
                }

                // Save the chapter as an individual HTML file.
                string chapterFilePath = Path.Combine(outputDir, $"Document_Chapter{chapterIndex}.html");
                chapterDoc.Save(chapterFilePath, SaveFormat.Html);

                // Move the outer loop index to the last processed paragraph.
                i = j - 1;
            }
        }

        // Verify that at least two split HTML files were generated.
        string[] htmlFiles = Directory.GetFiles(outputDir, "*.html");
        if (htmlFiles.Length < 2)
        {
            throw new Exception("Expected multiple split HTML files, but fewer were created.");
        }

        // Output the names of the generated files.
        foreach (string file in htmlFiles)
        {
            Console.WriteLine("Created: " + Path.GetFileName(file));
        }
    }
}
