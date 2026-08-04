using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply an emphasis mark to the builder's font.
        builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        builder.Write("East Asian text with emphasis mark.");

        // Retrieve the first Run in the document.
        Run run = (Run)doc.GetChildNodes(NodeType.Run, true)[0];
        Aspose.Words.EmphasisMark emphasis = run.Font.EmphasisMark;

        // Output the EmphasisMark value for debugging.
        Console.WriteLine($"EmphasisMark of the run: {emphasis}");

        // Save the document to a local file.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "EmphasisMarkDemo.docx");
        doc.Save(outputPath);
    }
}
