using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Define directories for input and output.
        string baseDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(baseDir, "Output");

        // Ensure the output directory exists.
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // Create a sample document with three sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - Paragraph 1");
        builder.Writeln("Section 3 - Paragraph 2");

        // Save the original document for reference.
        string sourcePath = Path.Combine(outputDir, "SourceDocument.docx");
        sourceDoc.Save(sourcePath);

        // Split the document by its sections.
        List<Document> splitDocuments = new List<Document>();

        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document splitDoc = new Document();

            // Remove any default nodes (the constructor adds a blank section).
            splitDoc.RemoveAllChildren();

            // Import the current section from the source document.
            Section importedSection = (Section)splitDoc.ImportNode(sourceDoc.Sections[i], true);

            // Append the imported section to the new document.
            splitDoc.AppendChild(importedSection);

            // Add the split document to the collection.
            splitDocuments.Add(splitDoc);

            // Save each split part.
            string splitPath = Path.Combine(outputDir, $"SplitPart_{i + 1}.docx");
            splitDoc.Save(splitPath);
        }

        // Validate that the expected number of split files were created.
        for (int i = 0; i < splitDocuments.Count; i++)
        {
            string expectedPath = Path.Combine(outputDir, $"SplitPart_{i + 1}.docx");
            if (!File.Exists(expectedPath))
                throw new FileNotFoundException($"Expected split file not found: {expectedPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Document split completed successfully.");
    }
}
