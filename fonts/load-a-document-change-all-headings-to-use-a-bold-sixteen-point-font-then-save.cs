using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Paths for temporary and final documents
        string samplePath = "sample.docx";
        string outputPath = "output.docx";

        // Create a sample document with headings
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Add Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Sample Heading 1");

        // Add normal paragraph
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph.");

        // Add Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Sample Heading 2");

        // Save the sample document
        sampleDoc.Save(samplePath, SaveFormat.Docx);

        // Load the document
        Document doc = new Document(samplePath);

        // Change all headings to bold 16-point font
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
            if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
            {
                foreach (Run run in para.Runs)
                {
                    // Use Aspose.Words.Font for text formatting
                    Aspose.Words.Font font = run.Font;
                    font.Bold = true;
                    font.Size = 16;

                    // Validation: ensure properties are set
                    if (!font.Bold || Math.Abs(font.Size - 16) > 0.01)
                    {
                        throw new InvalidOperationException("Font properties were not applied correctly.");
                    }
                }
            }
        }

        // Save the modified document
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the output file exists
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document processed and saved successfully.");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
