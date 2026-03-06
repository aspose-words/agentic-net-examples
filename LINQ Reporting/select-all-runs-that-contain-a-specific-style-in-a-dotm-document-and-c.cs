using System.Drawing;
using Aspose.Words;

class ChangeStyleRunsColor
{
    static void Main()
    {
        // Path to the source DOTM template.
        string inputPath = @"C:\Docs\Template.dotm";

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\Result.docx";

        // Load the DOTM document.
        Document doc = new Document(inputPath);

        // The style we want to target (by name).  You can also use
        // StyleIdentifier if you prefer to work with built‑in identifiers.
        const string targetStyleName = "MyCustomStyle";

        // Iterate through all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // If the run uses the specified style, change its font color.
            if (run.Font.StyleName == targetStyleName)
            {
                run.Font.Color = Color.Red;   // Set desired color.
            }

            // Example of using a built‑in style identifier instead of a name:
            // if (run.Font.StyleIdentifier == StyleIdentifier.IntenseEmphasis)
            // {
            //     run.Font.Color = Color.Red;
            // }
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
