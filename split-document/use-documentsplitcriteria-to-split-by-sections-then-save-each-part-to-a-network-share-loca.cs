using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            // Start a new section for each iteration (except the first, which already exists).
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Add some content that will be visible in the split output.
            builder.Writeln($"This is content of Section {i}.");
        }

        // Simulate a network share by creating a folder under the current directory.
        string networkSharePath = Path.Combine(Environment.CurrentDirectory, "NetworkShare");
        Directory.CreateDirectory(networkSharePath);

        // Split the document by its sections.
        int sectionIndex = 1;
        foreach (Section section in doc.Sections)
        {
            // Create a new empty document that will hold the single section.
            Document splitDoc = new Document();

            // Import the section from the source document into the new document.
            NodeImporter importer = new NodeImporter(doc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);

            // Remove the default empty section that a new Document contains.
            splitDoc.RemoveAllChildren();

            // Append the imported section as the sole section of the split document.
            splitDoc.AppendChild(importedSection);

            // Save each split part as an individual HTML file.
            string fileName = $"Section_{sectionIndex}.html";
            string filePath = Path.Combine(networkSharePath, fileName);
            splitDoc.Save(filePath, SaveFormat.Html);
            sectionIndex++;
        }

        // Validate that the expected number of split files were created.
        string[] splitFiles = Directory.GetFiles(networkSharePath, "*.html");
        if (splitFiles.Length != doc.Sections.Count)
            throw new InvalidOperationException(
                $"Expected {doc.Sections.Count} split files, but found {splitFiles.Length}.");

        // Output the paths of the generated files (no interactive prompts).
        foreach (string file in splitFiles)
        {
            Console.WriteLine($"Created split file: {file}");
        }
    }
}
