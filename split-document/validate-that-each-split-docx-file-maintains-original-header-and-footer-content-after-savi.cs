using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three sections, each having its own header and footer.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        int totalSections = 3;

        for (int i = 1; i <= totalSections; i++)
        {
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage); // start a new section

            // Header for the current section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write($"Header Sec{i}");

            // Footer for the current section.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write($"Footer Sec{i}");

            // Return to the main body and add some content.
            builder.MoveToDocumentEnd();
            builder.Writeln($"Content of section {i}");
        }

        // Save the source document (optional, just for inspection).
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Split the document by sections, preserving headers and footers.
        List<string> splitFiles = new List<string>();
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section originalSection = sourceDoc.Sections[i];

            // Create a new empty document without the default empty section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren();

            // Import the section (including its headers/footers) into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(originalSection, true);
            splitDoc.AppendChild(importedSection);

            // Save the split part.
            string splitPath = $"Split_{i + 1}.docx";
            splitDoc.Save(splitPath);
            splitFiles.Add(splitPath);
        }

        // Validate that each split file exists and retains its original header/footer text.
        foreach (string filePath in splitFiles)
        {
            if (!File.Exists(filePath))
                throw new Exception($"Expected split file not found: {filePath}");

            // Determine the section index from the file name (e.g., Split_2.docx -> 2).
            int index = int.Parse(Path.GetFileNameWithoutExtension(filePath).Split('_')[1]);

            Document splitDoc = new Document(filePath);
            HeaderFooter header = splitDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            HeaderFooter footer = splitDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            string headerText = header?.GetText().Trim() ?? string.Empty;
            string footerText = footer?.GetText().Trim() ?? string.Empty;

            string expectedHeader = $"Header Sec{index}";
            string expectedFooter = $"Footer Sec{index}";

            if (!headerText.Contains(expectedHeader))
                throw new Exception($"Header validation failed for {filePath}. Expected to contain \"{expectedHeader}\", but got \"{headerText}\".");

            if (!footerText.Contains(expectedFooter))
                throw new Exception($"Footer validation failed for {filePath}. Expected to contain \"{expectedFooter}\", but got \"{footerText}\".");
        }

        // If execution reaches this point, all validations passed.
        Console.WriteLine("All split documents have been validated successfully.");
    }
}
