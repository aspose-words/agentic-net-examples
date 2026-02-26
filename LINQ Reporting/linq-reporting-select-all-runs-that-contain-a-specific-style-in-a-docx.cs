using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document (lifecycle rule)
        Document doc = new Document("Input.docx");

        // The style name we want to target (case‑insensitive)
        const string targetStyleName = "MyCustomStyle";

        // Iterate over every Run node in the document
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // If the run's style matches the target style, change its font color
            if (string.Equals(run.Font.StyleName, targetStyleName, StringComparison.OrdinalIgnoreCase))
            {
                run.Font.Color = Color.Red; // Set desired color here
            }
        }

        // Save the modified document (lifecycle rule)
        doc.Save("Output.docx");
    }
}
