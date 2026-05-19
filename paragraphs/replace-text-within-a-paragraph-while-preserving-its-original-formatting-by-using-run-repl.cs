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

        // Apply some formatting to the builder's font.
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Font.Bold = true;
        builder.Font.Color = Color.Blue;

        // Write a paragraph that contains the text we will replace.
        builder.Writeln("Hello, this is a sample paragraph with placeholder text.");

        // Text to find and its replacement.
        const string oldText = "placeholder";
        const string newText = "replaced";

        // Iterate over all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // If the run contains the target substring, replace it.
            if (run.Text.Contains(oldText))
            {
                // Replace only the substring; the run's formatting remains unchanged.
                run.Text = run.Text.Replace(oldText, newText);
                break; // Assuming the text appears only once.
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
