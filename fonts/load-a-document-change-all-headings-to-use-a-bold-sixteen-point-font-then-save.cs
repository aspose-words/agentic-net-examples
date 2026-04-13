using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and the final output.
        const string inputPath = "Sample.docx";
        const string outputPath = "ModifiedHeadings.docx";

        // -----------------------------------------------------------------
        // Create a sample document that contains a few headings.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading Level 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading Level 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Normal paragraph.");

        // Save the sample document so we can load it later.
        sampleDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document from disk.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // -----------------------------------------------------------------
        // Change all headings to a bold 16‑point font.
        // -----------------------------------------------------------------
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph paragraph in paragraphs)
        {
            StyleIdentifier styleId = paragraph.ParagraphFormat.StyleIdentifier;

            // Check if the paragraph uses any heading style (Heading1 … Heading9).
            if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
            {
                // Retrieve the style object associated with this heading.
                Style headingStyle = doc.Styles[paragraph.ParagraphFormat.StyleName];

                // Apply the required font settings.
                headingStyle.Font.Size = 16;
                headingStyle.Font.Bold = true;

                // Validate that the properties were set correctly.
                if (headingStyle.Font.Size != 16 || headingStyle.Font.Bold != true)
                {
                    throw new InvalidOperationException("Failed to apply font settings to heading style.");
                }
            }
        }

        // -----------------------------------------------------------------
        // Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
