using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with two sections
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section content
        builder.Writeln("Section 1 - Paragraph 1");
        builder.Writeln("Section 1 - Paragraph 2");

        // Insert a section break to start a new section
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Second section content
        builder.Writeln("Section 2 - Paragraph 1");
        builder.Writeln("Section 2 - Paragraph 2");

        // Split the document by sections manually
        int sectionIndex = 1;
        foreach (Section srcSection in doc.Sections)
        {
            // Create a new empty document
            Document splitDoc = new Document();

            // Remove the default empty section that a new Document contains
            splitDoc.RemoveAllChildren();

            // Import the source section into the new document (deep clone with formatting)
            Section importedSection = (Section)splitDoc.ImportNode(srcSection, true, ImportFormatMode.KeepSourceFormatting);

            // Append the imported section as the sole section of the split document
            splitDoc.AppendChild(importedSection);

            // Save each split document as a separate HTML file
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection;

            string splitFilePath = Path.Combine(outputDir, $"Document_section{sectionIndex}.html");
            splitDoc.Save(splitFilePath, saveOptions);

            sectionIndex++;
        }

        // Validate that the expected number of HTML files were created
        string[] htmlFiles = Directory.GetFiles(outputDir, "*.html");
        if (htmlFiles.Length != doc.Sections.Count)
        {
            throw new InvalidOperationException(
                $"Expected {doc.Sections.Count} HTML files after splitting by sections, but found {htmlFiles.Length}.");
        }

        // Output the names of the generated files (no interactive prompt)
        foreach (string file in htmlFiles)
        {
            Console.WriteLine("Generated file: " + Path.GetFileName(file));
        }
    }
}
