using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Math;
using Aspose.Words.Drawing;

class FlagLongOfficeMath
{
    static void Main()
    {
        // Path to the folder that contains the DOTX template.
        string dataDir = @"C:\Docs";
        string inputFile = Path.Combine(dataDir, "Template.dotx");
        string outputFile = Path.Combine(dataDir, "Flagged.docx");

        // Load the DOTX document with OfficeMath conversion enabled.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.ConvertShapeToOfficeMath = true;
        Document doc = new Document(inputFile, loadOptions);

        // Length threshold for flagging equations.
        int lengthThreshold = 20;

        // Iterate through all OfficeMath nodes in the document.
        int mathCount = doc.GetChildNodes(NodeType.OfficeMath, true).Count;
        for (int i = 0; i < mathCount; i++)
        {
            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, i, true);
            string equationText = officeMath.GetText();

            // If the equation text exceeds the threshold, highlight its runs.
            if (equationText.Length > lengthThreshold)
            {
                foreach (Run run in officeMath.GetChildNodes(NodeType.Run, true))
                {
                    run.Font.HighlightColor = Color.Yellow;
                }
            }
        }

        // Save the modified document.
        doc.Save(outputFile);
    }
}
