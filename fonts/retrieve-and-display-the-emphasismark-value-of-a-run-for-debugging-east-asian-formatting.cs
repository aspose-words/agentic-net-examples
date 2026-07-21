using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply an emphasis mark to the font.
        builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        builder.Write("East Asian text with emphasis mark");

        // Save the document to a local file.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string docPath = Path.Combine(artifactsDir, "EmphasisMark.docx");
        doc.Save(docPath);

        // Retrieve the first run in the document.
        Run run = (Run)doc.GetChildNodes(NodeType.Run, true)[0];

        // Get the EmphasisMark value.
        Aspose.Words.EmphasisMark emphasis = run.Font.EmphasisMark;

        // Output the value for debugging.
        Console.WriteLine($"EmphasisMark value: {emphasis}");
    }
}
