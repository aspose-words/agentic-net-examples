using System;
using System.Drawing;
using Aspose.Words;

class ChangeRunStyleColor
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document("Input.docm");

        // Define the style name to look for.
        string targetStyleName = "MyCustomStyle";

        // Iterate through all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // Check if the run uses the specified style.
            if (run.Font.StyleName == targetStyleName)
            {
                // Change the font color of the run.
                run.Font.Color = Color.Red;
            }
        }

        // Save the modified document.
        doc.Save("Output.docm");
    }
}
