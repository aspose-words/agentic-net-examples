using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a run with an emphasis mark.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.EmphasisMark = Aspose.Words.EmphasisMark.OverSolidCircle; // Set emphasis.
        builder.Write("Emphasis text");

        // Save the document so that the output file exists.
        string outputPath = "EmphasisMark.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Retrieve the first Run in the document.
        Run run = (Run)doc.GetChildNodes(NodeType.Run, true)[0];

        // Get the EmphasisMark value from the Run's Font.
        Aspose.Words.EmphasisMark emphasis = run.Font.EmphasisMark;

        // Display the EmphasisMark value for debugging.
        Console.WriteLine($"EmphasisMark value: {emphasis}");
    }
}
