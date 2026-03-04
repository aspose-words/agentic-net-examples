using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input DOC file path.
            string inputPath = @"C:\Docs\InputDocument.docx";

            // Output DOC file path.
            string outputPath = @"C:\Docs\OutputDocument.docx";

            // The style name to search for (e.g., "Emphasis").
            string targetStyleName = "Emphasis";

            // Desired font color for the runs that use the target style.
            Color newFontColor = Color.Blue;

            // Load the document.
            Document doc = new Document(inputPath);

            // Iterate through all Run nodes in the document.
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                // Check if the run's character style matches the target style name.
                if (run.Font.StyleName == targetStyleName)
                {
                    // Change the font color of the matching run.
                    run.Font.Color = newFontColor;
                }
            }

            // Save the modified document.
            doc.Save(outputPath);
        }
    }
}
