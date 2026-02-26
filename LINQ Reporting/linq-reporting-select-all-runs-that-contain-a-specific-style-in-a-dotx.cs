using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOTX template document.
        Document doc = new Document("Template.dotx");

        // The style name we want to target (case‑insensitive).
        const string targetStyleName = "MyCustomStyle";

        // Iterate through every Run node in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // If the run's character style matches the target style, change its color.
            if (string.Equals(run.Font.StyleName, targetStyleName, StringComparison.OrdinalIgnoreCase))
            {
                run.Font.Color = Color.Red; // Set desired font color.
            }
        }

        // Save the modified document to a new file.
        doc.Save("Result.docx");
    }
}
