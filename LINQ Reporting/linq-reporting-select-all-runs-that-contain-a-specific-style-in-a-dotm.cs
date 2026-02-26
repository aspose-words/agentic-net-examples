using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load the DOTM template. The constructor handles the loading lifecycle.
        Document doc = new Document("Template.dotm");

        // Define the style name we want to target.
        const string targetStyleName = "MyCustomStyle";

        // Retrieve all Run nodes in the document.
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>();

        // Iterate through the runs and change the font color of those that use the target style.
        foreach (Run run in runs)
        {
            // Check by style name (case‑sensitive) or by built‑in identifier if needed.
            if (run.Font.StyleName == targetStyleName ||
                run.Font.StyleIdentifier == StyleIdentifier.IntenseEmphasis) // example of identifier check
            {
                // Change the font color to red.
                run.Font.Color = Color.Red;
            }
        }

        // Save the modified document. The Save method follows the prescribed lifecycle.
        doc.Save("Result.docx");
    }
}
