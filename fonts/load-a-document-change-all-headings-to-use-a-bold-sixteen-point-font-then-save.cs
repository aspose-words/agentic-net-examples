using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        string inputPath = "SampleInput.docx";
        string outputPath = "SampleOutput.docx";

        // -----------------------------------------------------------------
        // Create a sample document with headings if it does not already exist.
        // -----------------------------------------------------------------
        if (!File.Exists(inputPath))
        {
            Document createDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(createDoc);

            // Add a few headings of different levels.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading Level 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading Level 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Heading Level 3");

            // Add a normal paragraph to show that it will not be changed.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Regular paragraph text.");

            // Save the sample document.
            createDoc.Save(inputPath);
        }

        // --------------------------------------------------------------
        // Load the existing document, modify all headings, and save it.
        // --------------------------------------------------------------
        Document doc = new Document(inputPath);

        // Iterate over all paragraph nodes in the document.
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph style is any heading style.
            StyleIdentifier styleId = paragraph.ParagraphFormat.StyleIdentifier;
            bool isHeading = styleId == StyleIdentifier.Heading1 ||
                             styleId == StyleIdentifier.Heading2 ||
                             styleId == StyleIdentifier.Heading3 ||
                             styleId == StyleIdentifier.Heading4 ||
                             styleId == StyleIdentifier.Heading5 ||
                             styleId == StyleIdentifier.Heading6 ||
                             styleId == StyleIdentifier.Heading7 ||
                             styleId == StyleIdentifier.Heading8 ||
                             styleId == StyleIdentifier.Heading9;

            if (isHeading)
            {
                // Apply bold and 16‑point size to every run within the heading.
                foreach (Run run in paragraph.Runs)
                {
                    Aspose.Words.Font font = run.Font;
                    font.Bold = true;
                    font.Size = 16;
                }
            }
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
