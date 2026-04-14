using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add a run with an emphasis mark.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle; // Set emphasis.
        builder.Write("Emphasis text");

        // Retrieve the first run from the document.
        Run run = (Run)doc.GetChildNodes(NodeType.Run, true)[0];

        // Get the emphasis mark value from the run's font.
        Aspose.Words.EmphasisMark emphasis = run.Font.EmphasisMark;

        // Output the emphasis mark to the console for debugging.
        Console.WriteLine($"EmphasisMark of the run: {emphasis}");

        // Save the document to verify the formatting.
        doc.Save("EmphasisMark.docx");
    }
}
