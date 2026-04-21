using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with headings, each starting on a new page (new section)
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        void AddChapter(string title, int lineCount)
        {
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln(title);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            for (int i = 0; i < lineCount; i++)
                builder.Writeln($"Content line {i + 1} of {title}.");
        }

        AddChapter("Chapter 1", 30);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        AddChapter("Chapter 2", 30);
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        AddChapter("Chapter 3", 30);

        // Save the source document
        string sourcePath = Path.Combine(outputDir, "Original.docx");
        sourceDoc.Save(sourcePath);

        if (!File.Exists(sourcePath))
            throw new Exception("Failed to create the source document.");

        // Split the document: each section (starting with a heading) becomes a separate file
        Document src = new Document(sourcePath);
        for (int i = 0; i < src.Sections.Count; i++)
        {
            // Create a new document and import the current section
            Document part = new Document();
            part.Sections.Clear(); // Remove the default empty section

            Section importedSection = (Section)part.ImportNode(src.Sections[i], true);
            part.Sections.Add(importedSection);

            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            part.Save(partPath);

            if (!File.Exists(partPath))
                throw new Exception($"Failed to create split part: {partPath}");
        }

        Console.WriteLine($"Document split completed. Files are located in: {outputDir}");
    }
}
