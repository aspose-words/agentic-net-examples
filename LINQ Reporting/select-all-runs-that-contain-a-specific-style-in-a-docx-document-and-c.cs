using System.Drawing;
using Aspose.Words;

class ChangeStyleRunsColor
{
    static void Main()
    {
        // Path to the source DOCX file.
        const string inputPath = @"input.docx";

        // Path where the modified document will be saved.
        const string outputPath = @"output.docx";

        // The name of the style whose runs we want to recolor.
        const string targetStyleName = "Emphasis";

        // Load the document (lifecycle rule: load).
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

        // Save the modified document (lifecycle rule: save).
        doc.Save(outputPath);
    }
}
