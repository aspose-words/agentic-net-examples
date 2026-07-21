using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define input and output file paths.
        string inputPath = "input.docx";
        string outputPath = "output.docx";

        // If the input file does not exist, create a sample document with headings.
        if (!File.Exists(inputPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Create headings 1‑3.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Sample Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Sample Heading 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Sample Heading 3");

            // Normal paragraph.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Regular paragraph text.");

            // Save the sample document to the expected input location.
            sampleDoc.Save(inputPath);
        }

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Iterate through all paragraphs and apply bold 16‑point font to headings.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier styleId = paragraph.ParagraphFormat.StyleIdentifier;
            if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
            {
                foreach (Run run in paragraph.Runs)
                {
                    Aspose.Words.Font font = run.Font;
                    font.Bold = true;
                    font.Size = 16;

                    // Validate the formatting.
                    if (!font.Bold || font.Size != 16)
                        throw new InvalidOperationException("Failed to set heading font properties.");
                }
            }
        }

        // Save the modified document.
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
