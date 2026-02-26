using System;
using Aspose.Words;
using Aspose.Words.Math;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source TXT file that may contain OfficeMath objects.
        string inputPath = @"C:\Docs\Input.txt";

        // Path where the resulting TXT file will be saved.
        string outputPath = @"C:\Docs\Output.txt";

        // Load the TXT document with default load options.
        // TxtLoadOptions allows additional control if needed (e.g., ConvertShapeToOfficeMath).
        Document doc = new Document(inputPath, new TxtLoadOptions());

        // Iterate over all OfficeMath nodes in the document.
        // Setting DisplayType to Display forces each equation to be on its own line.
        foreach (OfficeMath officeMath in doc.GetChildNodes(NodeType.OfficeMath, true))
        {
            officeMath.DisplayType = OfficeMathDisplayType.Display;
        }

        // Prepare save options for TXT format.
        // OfficeMathExportMode = Text (default) ensures equations are exported as plain text.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            OfficeMathExportMode = TxtOfficeMathExportMode.Text
        };

        // Save the modified document back to TXT.
        doc.Save(outputPath, saveOptions);
    }
}
