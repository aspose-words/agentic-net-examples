using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Path to the source DOCM document.
        string inputPath = @"C:\Docs\input.docm";

        // Load the document using the provided constructor.
        Document doc = new Document(inputPath);

        // Define the style to search for. You can use either the style name or the style identifier.
        // Example using the style name "Emphasis":
        string targetStyleName = "Emphasis";

        // Example using the style identifier (optional):
        // StyleIdentifier targetStyleId = StyleIdentifier.Emphasis;

        // Select all Run nodes in the document.
        var runs = doc.GetChildNodes(NodeType.Run, true).Cast<Run>();

        // Filter runs that have the specified style and change their font color.
        foreach (Run run in runs.Where(r => 
                     string.Equals(r.Font.StyleName, targetStyleName, StringComparison.OrdinalIgnoreCase)
                     // Uncomment the following line to filter by style identifier instead:
                     // || r.Font.StyleIdentifier == targetStyleId
                 ))
        {
            // Change the font color to red (choose any color you need).
            run.Font.Color = Color.Red;
        }

        // Save the modified document using the provided Save method.
        string outputPath = @"C:\Docs\output.docm";
        doc.Save(outputPath);
    }
}
