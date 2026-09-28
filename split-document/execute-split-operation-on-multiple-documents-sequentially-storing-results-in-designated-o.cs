using System;
using System.IO;
using Aspose.Words;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define folders for input and output.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputDocs");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample source documents.
        CreateSampleDocument(Path.Combine(inputFolder, "Doc1.docx"), 2);
        CreateSampleDocument(Path.Combine(inputFolder, "Doc2.docx"), 3);

        // Process each document: split by sections.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document source = new Document(docPath);
            string docNameWithoutExt = Path.GetFileNameWithoutExtension(docPath);

            int sectionIndex = 1;
            foreach (Section section in source.Sections)
            {
                // Create a new document for the current section.
                Document splitDoc = new Document();
                splitDoc.RemoveAllChildren(); // Remove the default empty section.

                // Import the section from the source document.
                Section importedSection = (Section)splitDoc.ImportNode(section, true, ImportFormatMode.KeepSourceFormatting);
                splitDoc.AppendChild(importedSection);

                // Build output file name.
                string outFileName = $"{docNameWithoutExt}_Section{sectionIndex}.docx";
                string outPath = Path.Combine(outputFolder, outFileName);

                // Save the split document.
                splitDoc.Save(outPath);
                sectionIndex++;
            }
        }

        // Validation: ensure each expected split file exists.
        foreach (string docPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document source = new Document(docPath);
            string docNameWithoutExt = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docPath));

            int expectedCount = source.Sections.Count;
            for (int i = 1; i <= expectedCount; i++)
            {
                string expectedPath = $"{docNameWithoutExt}_Section{i}.docx";
                if (!File.Exists(expectedPath))
                {
                    throw new FileNotFoundException($"Expected split file not found: {expectedPath}");
                }
            }
        }

        // Indicate successful completion.
        Console.WriteLine("Document splitting completed successfully.");
    }

    // Helper method to create a sample document with a given number of sections.
    private static void CreateSampleDocument(string filePath, int sectionsCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= sectionsCount; i++)
        {
            // Add a new section for each iteration except the first (the document already has one).
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Add a heading to identify the section.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Section {i} Heading");

            // Add some body text.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln($"This is the content of section {i}.");
        }

        // Save the sample document.
        doc.Save(filePath);
    }
}
