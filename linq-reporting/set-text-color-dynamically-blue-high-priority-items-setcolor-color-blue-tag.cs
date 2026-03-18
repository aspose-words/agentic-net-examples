using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Iterate over all runs in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // Look for the custom tag <<setColor color="blue">>.
            if (run.Text.Contains("<<setColor") && run.Text.Contains("color=\"blue\""))
            {
                // Apply blue color to the run's font.
                run.Font.Color = Color.Blue;

                // Optionally remove the tag text itself.
                // Example: replace the whole tag with the inner content (if any).
                // Here we simply clear the tag text, assuming the actual content follows it.
                // Adjust as needed for your specific markup.
                run.Text = run.Text.Replace("<<setColor color=\"blue\">>", string.Empty);
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
