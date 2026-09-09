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
        builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle;
        builder.Write("East Asian text with emphasis");

        // Save the document so that an output file exists.
        const string outputPath = "EmphasisMark.docx";
        doc.Save(outputPath);

        // Retrieve the first Run in the document.
        Run run = (Run)doc.GetChild(NodeType.Run, 0, true);

        // Get the EmphasisMark value from the Run's Font.
        Aspose.Words.EmphasisMark emphasis = run.Font.EmphasisMark;

        // Display the EmphasisMark value.
        Console.WriteLine($"EmphasisMark value: {emphasis}");
    }
}
