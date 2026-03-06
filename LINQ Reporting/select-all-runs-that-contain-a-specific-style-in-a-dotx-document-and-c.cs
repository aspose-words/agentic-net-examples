using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOTX template document.
        Document doc = new Document("Template.dotx");

        // The style name we want to target.
        const string targetStyleName = "MyCustomStyle";

        // Iterate over every Run node in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // If the run uses the specified style, change its font color.
            if (run.Font.StyleName == targetStyleName)
            {
                run.Font.Color = Color.Red; // Set desired color.
            }
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
