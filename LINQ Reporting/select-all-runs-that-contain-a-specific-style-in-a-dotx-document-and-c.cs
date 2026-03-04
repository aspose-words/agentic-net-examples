using System;
using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source DOTX template.
        string inputPath = "Template.dotx";

        // Path where the modified document will be saved.
        string outputPath = "Result.docx";

        // The name of the style to search for.
        string targetStyleName = "MyStyle";

        // Load the DOTX document.
        Document doc = new Document(inputPath);

        // Iterate through all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // Check if the run uses the specified style.
            if (run.Font.StyleName == targetStyleName)
            {
                // Change the font color of the run.
                run.Font.Color = Color.Blue;
            }
        }

        // Save the updated document.
        doc.Save(outputPath);
    }
}
