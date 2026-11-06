using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks; // Used for DocumentBuilder

public class Program
{
    public static void Main()
    {
        // Base directories for input and output.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");

        // Ensure a clean environment.
        if (Directory.Exists(inputDir)) Directory.Delete(inputDir, true);
        if (Directory.Exists(outputDir)) Directory.Delete(outputDir, true);
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample source documents.
        int sourceDocCount = 2;
        for (int i = 1; i <= sourceDocCount; i++)
        {
            Document srcDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(srcDoc);

            // Section 1
            builder.Writeln($"Document {i} - Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Section 2
            builder.Writeln($"Document {i} - Section 2");

            string srcPath = Path.Combine(inputDir, $"Doc{i}.docx");
            srcDoc.Save(srcPath);
        }

        // Process each source document: split by sections.
        string[] sourceFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string srcFilePath in sourceFiles)
        {
            // Load source document.
            Document sourceDoc = new Document(srcFilePath);
            string srcFileName = Path.GetFileNameWithoutExtension(srcFilePath);
            string docOutputDir = Path.Combine(outputDir, srcFileName);
            Directory.CreateDirectory(docOutputDir);

            int sectionIndex = 1;
            foreach (Section srcSection in sourceDoc.Sections)
            {
                // Create a new empty document.
                Document splitDoc = new Document();
                splitDoc.RemoveAllChildren(); // Remove the default empty section.

                // Import the source section into the new document.
                NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
                Section importedSection = (Section)importer.ImportNode(srcSection, true);
                splitDoc.AppendChild(importedSection);

                // Save the split document.
                string outPath = Path.Combine(docOutputDir, $"Section{sectionIndex}.docx");
                splitDoc.Save(outPath);
                sectionIndex++;
            }
        }

        // Validation: ensure each expected split file exists.
        foreach (string srcFilePath in sourceFiles)
        {
            string srcFileName = Path.GetFileNameWithoutExtension(srcFilePath);
            string docOutputDir = Path.Combine(outputDir, srcFileName);
            // Expecting two sections per source document.
            for (int s = 1; s <= 2; s++)
            {
                string expectedPath = Path.Combine(docOutputDir, $"Section{s}.docx");
                if (!File.Exists(expectedPath))
                {
                    throw new FileNotFoundException($"Expected split file not found: {expectedPath}");
                }
            }
        }

        // Indicate successful completion (no interactive prompts).
        Console.WriteLine("Document splitting completed successfully.");
    }
}
