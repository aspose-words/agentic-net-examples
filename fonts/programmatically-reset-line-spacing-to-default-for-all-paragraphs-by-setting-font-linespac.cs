using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Use DocumentBuilder to add sample paragraphs with custom line spacing.
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 3; i++)
        {
            // Set a non‑default line spacing (e.g., 12 points).
            builder.ParagraphFormat.LineSpacing = 12;
            builder.Writeln($"Paragraph {i + 1} with custom line spacing.");
        }

        // Reset line spacing to default (0) for all paragraphs.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            para.ParagraphFormat.LineSpacing = 0;

            // Validate that the line spacing was reset.
            if (para.ParagraphFormat.LineSpacing != 0)
            {
                throw new Exception("Line spacing reset failed.");
            }
        }

        // Save the document.
        string outputPath = "ResetLineSpacing.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
