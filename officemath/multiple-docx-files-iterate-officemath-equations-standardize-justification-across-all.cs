using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Math;

namespace OfficeMathJustificationStandardizer
{
    public class Program
    {
        public static void Main()
        {
            // Use folders relative to the executable location.
            string baseDir = AppContext.BaseDirectory;
            string sourceFolder = Path.Combine(baseDir, "Input");
            string destinationFolder = Path.Combine(baseDir, "Output");

            // Ensure both folders exist.
            Directory.CreateDirectory(sourceFolder);
            Directory.CreateDirectory(destinationFolder);

            // Target justification to apply to every OfficeMath equation.
            OfficeMathJustification targetJustification = OfficeMathJustification.Center;

            // Gather all DOCX files from the source folder.
            string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx");

            if (docxFiles.Length == 0)
            {
                Console.WriteLine($"No DOCX files found in '{sourceFolder}'. Place files there and rerun.");
                return;
            }

            // Process each document.
            foreach (string filePath in docxFiles)
            {
                StandardizeOfficeMathJustification(filePath, destinationFolder, targetJustification);
                Console.WriteLine($"Processed: {Path.GetFileName(filePath)}");
            }

            Console.WriteLine($"All files saved to '{destinationFolder}'.");
        }

        private static void StandardizeOfficeMathJustification(string inputFilePath, string outputFolder, OfficeMathJustification justification)
        {
            // Load the document.
            Document doc = new Document(inputFilePath);

            // Retrieve all OfficeMath nodes (deep search).
            NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

            // Apply justification to each OfficeMath node.
            foreach (OfficeMath officeMath in officeMathNodes.OfType<OfficeMath>())
            {
                if (officeMath.DisplayType == OfficeMathDisplayType.Inline)
                    officeMath.DisplayType = OfficeMathDisplayType.Display;

                officeMath.Justification = justification;
            }

            // Save the modified document.
            string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(inputFilePath));
            doc.Save(outputFilePath);
        }
    }
}
