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

        // Apply an emphasis mark to the text.
        builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        builder.Write("East Asian text with emphasis");

        // Retrieve the first Run in the document.
        Run run = (Run)doc.GetChildNodes(NodeType.Run, true)[0];

        // Get the EmphasisMark value from the Run's Font.
        Aspose.Words.EmphasisMark emphasis = run.Font.EmphasisMark;

        // Output the EmphasisMark value for debugging.
        Console.WriteLine($"EmphasisMark value: {emphasis}");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmphasisMark.docx");
        doc.Save(outputPath);
    }
}
