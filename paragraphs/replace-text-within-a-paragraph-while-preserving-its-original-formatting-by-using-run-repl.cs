using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a paragraph that contains several runs with different formatting.
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Font.Color = Color.Black;
        builder.Write("Hello ");

        builder.Font.Bold = true;
        builder.Write("World");

        builder.Font.Bold = false;
        builder.Write("! This is a sample paragraph with the word ");

        builder.Font.Italic = true;
        builder.Write("Aspose.Words");

        builder.Font.Italic = false;
        builder.Write(" library.");

        // Text to be replaced and its replacement.
        const string oldText = "Aspose.Words";
        const string newText = "Aspose.Words.NET";

        // Iterate over all runs in the document and replace the target text.
        // Changing the Run.Text property preserves the original formatting of each run.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text.Contains(oldText))
                run.Text = run.Text.Replace(oldText, newText);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
