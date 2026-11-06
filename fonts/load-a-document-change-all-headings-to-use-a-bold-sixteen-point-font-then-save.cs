using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document containing headings.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        // Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");

        // Normal paragraph
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Some regular text.");

        // Save the sample document.
        string inputPath = "sample.docx";
        sampleDoc.Save(inputPath);

        // Load the document.
        Document doc = new Document(inputPath);

        // Change all headings to bold 16‑point font.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier styleId = paragraph.ParagraphFormat.StyleIdentifier;
            if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
            {
                // Apply formatting to each Run in the paragraph.
                foreach (Run run in paragraph.Runs)
                {
                    run.Font.Bold = true;
                    run.Font.Size = 16;

                    // Validate the changes.
                    if (!run.Font.Bold || run.Font.Size != 16)
                    {
                        throw new InvalidOperationException("Failed to apply font formatting.");
                    }
                }
            }
        }

        // Save the modified document.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output file was not created.", outputPath);
        }
    }
}
