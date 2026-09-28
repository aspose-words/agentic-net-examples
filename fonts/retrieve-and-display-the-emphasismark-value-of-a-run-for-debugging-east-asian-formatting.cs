using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set East Asian font and an emphasis mark.
        builder.Font.Name = "MS Mincho";

        // Use a valid EmphasisMark value that exists in the current Aspose.Words version.
        builder.Font.EmphasisMark = EmphasisMark.None; // Change to a different value if needed.

        builder.Writeln("Sample text with emphasis mark.");

        // Retrieve the first Run in the document.
        Run run = (Run)doc.GetChildNodes(NodeType.Run, true)[0];
        EmphasisMark emphasis = run.Font.EmphasisMark;

        // Display the EmphasisMark value.
        Console.WriteLine($"EmphasisMark: {emphasis}");

        // Save the document to verify output.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Confirm the file was saved.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved to {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
