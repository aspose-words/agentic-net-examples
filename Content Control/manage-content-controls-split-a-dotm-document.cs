using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ManageContentControls
{
    class Program
    {
        static void Main(string[] args)
        {
            // Validate arguments.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: ManageContentControls <sourceDotmPath> <outputFolder>");
                return;
            }

            string sourceDotmPath = args[0];
            string outputFolder = args[1];

            var splitter = new DotmSplitter();
            splitter.SplitDotmByContentControls(sourceDotmPath, outputFolder);
        }
    }

    public class DotmSplitter
    {
        /// <summary>
        /// Splits a DOTM (macro‑enabled template) into separate documents.
        /// Each top‑level content control (StructuredDocumentTag) becomes its own DOCX file.
        /// </summary>
        /// <param name="sourceDotmPath">Full path to the source .dotm file.</param>
        /// <param name="outputFolder">Folder where the split documents will be saved.</param>
        public void SplitDotmByContentControls(string sourceDotmPath, string outputFolder)
        {
            // Load the source DOTM document.
            Document sourceDoc = new Document(sourceDotmPath);

            // Ensure the output folder exists.
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Retrieve all StructuredDocumentTag nodes (content controls).
            // The second argument 'true' makes the search recursive.
            NodeCollection sdtNodes = sourceDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            int partIndex = 1;
            foreach (StructuredDocumentTag sdt in sdtNodes)
            {
                // Consider only top‑level content controls (direct children of the document body).
                if (sdt.ParentNode?.NodeType != NodeType.Body)
                    continue;

                // Create a new blank document that will hold the extracted part.
                Document partDoc = new Document();
                partDoc.EnsureMinimum(); // Guarantees at least one section and one paragraph.

                // Import the content control (including its children) into the new document.
                NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);
                Node importedNode = importer.ImportNode(sdt, true);

                // Append the imported node to the body of the new document.
                partDoc.FirstSection.Body.AppendChild(importedNode);

                // Build a file name for the part.
                string partFileName = Path.Combine(outputFolder, $"Part_{partIndex}.docx");

                // Save the part.
                partDoc.Save(partFileName);

                Console.WriteLine($"Saved: {partFileName}");
                partIndex++;
            }
        }
    }
}
