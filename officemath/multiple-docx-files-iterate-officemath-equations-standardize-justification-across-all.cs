using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Math;

namespace OfficeMathJustificationStandardizer
{
    public class Program
    {
        // Entry point for demonstration purposes.
        public static void Main()
        {
            // Example input: all DOCX files in a source directory.
            string sourceFolder = @"C:\Docs\Input";
            string destinationFolder = @"C:\Docs\Output";

            // Ensure the output folder exists.
            Directory.CreateDirectory(destinationFolder);

            // Target justification to apply to every OfficeMath equation.
            OfficeMathJustification targetJustification = OfficeMathJustification.Center;

            // Gather all DOCX files from the source folder.
            string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx");

            // Process each document.
            foreach (string filePath in docxFiles)
            {
                StandardizeOfficeMathJustification(filePath, destinationFolder, targetJustification);
            }
        }

        /// <summary>
        /// Loads a DOCX file, sets a uniform justification for all OfficeMath objects,
        /// and saves the modified document to the specified output folder.
        /// </summary>
        /// <param name="inputFilePath">Full path to the source DOCX file.</param>
        /// <param name="outputFolder">Folder where the processed file will be saved.</param>
        /// <param name="justification">Desired OfficeMathJustification value.</param>
        private static void StandardizeOfficeMathJustification(string inputFilePath, string outputFolder, OfficeMathJustification justification)
        {
            // Load the document using the Document(string) constructor.
            Document doc = new Document(inputFilePath);

            // Retrieve all OfficeMath nodes in the document (deep search).
            NodeCollection officeMathNodes = doc.GetChildNodes(NodeType.OfficeMath, true);

            // Iterate through each OfficeMath node and apply the justification.
            foreach (OfficeMath officeMath in officeMathNodes.OfType<OfficeMath>())
            {
                // Justification cannot be set while the equation is inline.
                // Ensure the display type is not Inline before assigning justification.
                if (officeMath.DisplayType == OfficeMathDisplayType.Inline)
                {
                    officeMath.DisplayType = OfficeMathDisplayType.Display;
                }

                // Apply the requested justification.
                officeMath.Justification = justification;
            }

            // Build the output file path (preserve original file name).
            string outputFilePath = Path.Combine(outputFolder, Path.GetFileName(inputFilePath));

            // Save the modified document using the Save(string) method.
            doc.Save(outputFilePath);
        }
    }
}
