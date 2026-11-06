using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with formatted text.
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Font.Bold = true;
        builder.Writeln("Hello world! This is a sample paragraph.");

        // Text to replace.
        string oldText = "world";
        string newText = "universe";

        // Replace the text while preserving the original formatting.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            foreach (Run run in para.Runs)
            {
                if (run.Text.Contains(oldText))
                {
                    run.Text = run.Text.Replace(oldText, newText);
                }
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
