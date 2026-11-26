using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample source documents.
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.docx"), 3);
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.docx"), 2);

        // Process each document sequentially.
        foreach (string sourcePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document sourceDoc = new Document(sourcePath);
            string docName = Path.GetFileNameWithoutExtension(sourcePath);
            string docOutputDir = Path.Combine(outputDir, docName);
            Directory.CreateDirectory(docOutputDir);

            int sectionCount = sourceDoc.Sections.Count;
            for (int i = 0; i < sectionCount; i++)
            {
                Section section = sourceDoc.Sections[i];
                Document splitDoc = new Document();

                // Import the section into the new document.
                NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
                Section importedSection = (Section)importer.ImportNode(section, true);
                splitDoc.AppendChild(importedSection);

                string splitPath = Path.Combine(docOutputDir, $"Section_{i + 1}.docx");
                splitDoc.Save(splitPath);

                // Validate that the split file was created.
                if (!File.Exists(splitPath))
                    throw new InvalidOperationException($"Failed to create split file: {splitPath}");
            }

            // Verify the expected number of split files.
            string[] splitFiles = Directory.GetFiles(docOutputDir, "Section_*.docx");
            if (splitFiles.Length != sectionCount)
                throw new InvalidOperationException($"Expected {sectionCount} split files for {docName}, but found {splitFiles.Length}.");
        }
    }

    // Helper method to create a sample document with a given number of sections.
    private static void CreateSampleDocument(string path, int sections)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < sections; i++)
        {
            builder.Writeln($"This is content of Section {i + 1} in document {Path.GetFileName(path)}.");
            if (i < sections - 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        doc.Save(path);
    }
}
