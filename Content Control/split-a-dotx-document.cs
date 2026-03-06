using System;
using System.IO;
using Aspose.Words;

namespace SplitDotxExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOTX template.
            string sourcePath = @"C:\Docs\Template.dotx";

            // Directory where the split parts will be saved.
            string outputDir = @"C:\Docs\SplitParts";

            // Ensure the output directory exists.
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the DOTX document.
            Document sourceDoc = new Document(sourcePath);

            // Iterate through each section in the source document.
            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Create a new empty document.
                Document partDoc = new Document();

                // Import the current section into the new document.
                // ImportNode clones the node and adapts it to the target document.
                Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true);

                // Append the imported section to the new document.
                partDoc.AppendChild(importedSection);

                // Build a filename for the part (e.g., Template_part_1.dotx).
                string partFileName = Path.Combine(outputDir,
                    $"Template_part_{i + 1}{Path.GetExtension(sourcePath)}");

                // Save the part as a DOTX file.
                partDoc.Save(partFileName);
            }

            Console.WriteLine("Document split into parts successfully.");
        }
    }
}
