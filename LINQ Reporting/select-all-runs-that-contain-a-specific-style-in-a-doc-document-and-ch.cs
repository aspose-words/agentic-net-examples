using System.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // The style we want to target (change as needed).
        const string targetStyleName = "Emphasis";

        // The new font color to apply.
        Color newColor = Color.Red;

        // Iterate over every Run in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // If the run uses the specified style, change its font color.
            if (run.Font.StyleName == targetStyleName)
            {
                run.Font.Color = newColor;
            }

            // Example of checking by built‑in identifier instead of name:
            // if (run.Font.StyleIdentifier == StyleIdentifier.IntenseEmphasis)
            // {
            //     run.Font.Color = newColor;
            // }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
