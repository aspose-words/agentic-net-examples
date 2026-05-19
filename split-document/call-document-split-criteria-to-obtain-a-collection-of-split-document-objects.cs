using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output folder and ensure it exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three sections
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is the content of section {i}.");

            // Insert a section break after each section except the last one
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document (optional, for inspection)
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // Split the document by sections
        for (int idx = 0; idx < sourceDoc.Sections.Count; idx++)
        {
            // Clone the current section
            Section clonedSection = sourceDoc.Sections[idx].Clone();

            // Create a new document and import the cloned section into it
            Document splitDoc = new Document();

            // ImportNode requires an ImportFormatMode enum as the third argument
            Section importedSection = (Section)splitDoc.ImportNode(
                clonedSection,
                true, // Import child nodes
                ImportFormatMode.KeepSourceFormatting);

            // Append the imported section to the new document
            splitDoc.AppendChild(importedSection);

            // Ensure the document has the minimal required structure
            splitDoc.EnsureMinimum();

            // Save the split document
            string splitPath = Path.Combine(outputDir, $"Section_{idx + 1}.docx");
            splitDoc.Save(splitPath);
        }

        // Validate that all split files were created
        for (int i = 1; i <= sourceDoc.Sections.Count; i++)
        {
            string expectedPath = Path.Combine(outputDir, $"Section_{i}.docx");
            if (!File.Exists(expectedPath))
                throw new FileNotFoundException($"Expected split file not found: {expectedPath}");
        }

        // Indicate successful completion
        Console.WriteLine("Document split into sections successfully. Files are located at:");
        Console.WriteLine(outputDir);
    }
}
