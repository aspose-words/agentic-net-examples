using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");

        // Prepare output folder.
        string outputDir = "SplitOutput";
        Directory.CreateDirectory(outputDir);

        // Save the original document (optional, just for reference).
        string originalPath = Path.Combine(outputDir, "OriginalDocument.html");
        doc.Save(originalPath, SaveFormat.Html);

        // Split the document by sections.
        for (int i = 0; i < doc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document splitDoc = new Document();
            // Remove the default empty section that a new Document contains.
            splitDoc.RemoveAllChildren();

            // Import the current section from the source document.
            Section importedSection = (Section)splitDoc.ImportNode(doc.Sections[i], true);
            splitDoc.AppendChild(importedSection);

            // Save each split section as a separate HTML file.
            string splitPath = Path.Combine(outputDir, $"SplitDocument_section{i + 1}.html");
            splitDoc.Save(splitPath, SaveFormat.Html);
        }

        // Validate that the original file exists.
        if (!File.Exists(originalPath))
            throw new Exception("Original output file was not created.");

        // Validate that at least one section split file was generated.
        string[] splitFiles = Directory.GetFiles(outputDir, "SplitDocument_section*.html");
        if (splitFiles.Length == 0)
            throw new Exception("No section split files were created.");

        // Indicate successful execution.
        Console.WriteLine("Document split into sections successfully.");
    }
}
