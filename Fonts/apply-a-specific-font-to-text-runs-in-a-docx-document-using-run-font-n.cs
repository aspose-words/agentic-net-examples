using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class ApplyFontToRuns
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = "input.docx";

        // Path where the modified DOCX will be saved.
        string outputPath = "output.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Iterate over all Run nodes in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            // Apply the desired font name to each run.
            run.Font.Name = "Arial";
        }

        // Save the modified document.
        doc.Save(outputPath);
    }
}
