using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Define input and output file paths.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputPath = Path.Combine(dataDir, "Input.docx");
        string outputPath = Path.Combine(dataDir, "Output.docx");

        // Ensure the Data directory exists.
        if (!Directory.Exists(dataDir))
            Directory.CreateDirectory(dataDir);

        // If the input file does not exist, create a simple sample document with headings.
        if (!File.Exists(inputPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Sample Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Sample Heading 2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Normal paragraph text.");

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
                }
            }
        }

        // Save the modified document.
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (File.Exists(outputPath))
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        else
            Console.WriteLine("Failed to save the document.");
    }
}
