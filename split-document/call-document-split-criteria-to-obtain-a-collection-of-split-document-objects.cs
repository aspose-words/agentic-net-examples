using System;
using System.IO;
using Aspose.Words;

namespace SplitDocumentExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a sample document containing three sections.
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            builder.Writeln("Content of Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.Writeln("Content of Section 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.Writeln("Content of Section 3");

            // Split the source document by its sections.
            for (int i = 0; i < sourceDoc.Sections.Count; i++)
            {
                // Create a new empty document for the current part.
                Document splitDoc = new Document();
                // Remove the default empty section that a new Document contains.
                splitDoc.RemoveAllChildren();

                // Import the section from the source document into the new document.
                // ImportNode clones the node and re‑parents it to the target document.
                Section importedSection = (Section)splitDoc.ImportNode(sourceDoc.Sections[i], true);
                splitDoc.AppendChild(importedSection);

                // Save the split part.
                string partPath = Path.Combine(outputDir, $"SplitPart_{i + 1}.docx");
                splitDoc.Save(partPath);

                // Verify that the file was created.
                if (!File.Exists(partPath))
                    throw new InvalidOperationException($"Failed to create split document: {partPath}");
            }

            Console.WriteLine($"Document split into {sourceDoc.Sections.Count} parts. Files saved to: {outputDir}");
        }
    }
}
