using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

namespace SplitDocumentExample
{
    public class Program
    {
        public static void Main()
        {
            // Define paths for the sample source document and the split output folder.
            string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
            Directory.CreateDirectory(dataDir);
            string sourcePath = Path.Combine(dataDir, "Source.docx");
            string outputDir = Path.Combine(dataDir, "SplitSections");
            Directory.CreateDirectory(outputDir);

            // -------------------------------------------------
            // 1. Create a sample DOCX document with multiple sections.
            // -------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sourceDoc);

            // Section 1
            builder.Writeln("This is content of Section 1.");
            // Insert a section break to start a new section on a new page.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 2
            builder.Writeln("This is content of Section 2.");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 3
            builder.Writeln("This is content of Section 3.");

            // Save the sample source document.
            sourceDoc.Save(sourcePath);

            // -------------------------------------------------
            // 2. Load the DOCX source document using the Document class.
            // -------------------------------------------------
            Document loadedDoc = new Document(sourcePath);

            // -------------------------------------------------
            // 3. Split the loaded document by its sections.
            // -------------------------------------------------
            int sectionCount = loadedDoc.Sections.Count;
            for (int i = 0; i < sectionCount; i++)
            {
                // Create a new empty document.
                Document splitDoc = new Document();

                // Remove the default empty section that Aspose.Words creates.
                splitDoc.RemoveAllChildren();

                // Clone the current section from the source document.
                Section clonedSection = loadedDoc.Sections[i].Clone();

                // Import the cloned section into the new document to preserve styles and resources.
                Node importedSection = splitDoc.ImportNode(clonedSection, true, ImportFormatMode.KeepSourceFormatting);
                splitDoc.AppendChild(importedSection);

                // Save the split document.
                string splitPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");
                splitDoc.Save(splitPath);
            }

            // -------------------------------------------------
            // 4. Validate that the expected split output files exist.
            // -------------------------------------------------
            for (int i = 0; i < sectionCount; i++)
            {
                string expectedPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");
                if (!File.Exists(expectedPath))
                {
                    throw new FileNotFoundException($"Expected split file not found: {expectedPath}");
                }
            }

            // Indicate successful completion.
            Console.WriteLine($"Document split into {sectionCount} sections successfully.");
        }
    }
}
