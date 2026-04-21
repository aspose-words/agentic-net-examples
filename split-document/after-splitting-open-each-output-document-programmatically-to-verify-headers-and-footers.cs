using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample source document with three sections, each having its own header and footer.
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 3; i++)
        {
            // Set header for the current section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write($"Header {i}");

            // Set footer for the current section.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write($"Footer {i}");

            // Return to the main body and write some content.
            builder.MoveToDocumentEnd();
            builder.Writeln($"Content of section {i}");

            // Insert a section break after each section except the last one.
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document.
        sourceDoc.Save(sourcePath);

        // Split the source document by sections, preserving headers and footers.
        List<string> splitFiles = new List<string>();
        for (int idx = 0; idx < sourceDoc.Sections.Count; idx++)
        {
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // Ensure the document is empty.

            // Import the section from the source document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sourceDoc.Sections[idx], true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(outputDir, $"Section_{idx + 1}.docx");
            splitDoc.Save(splitPath);
            splitFiles.Add(splitPath);
        }

        // Verify that each split document exists and that its header/footer contain the expected text.
        foreach (string filePath in splitFiles)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Expected split file not found: {filePath}");

            Document part = new Document(filePath);
            HeaderFooter header = part.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            HeaderFooter footer = part.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            string headerText = header?.GetText().Trim() ?? string.Empty;
            string footerText = footer?.GetText().Trim() ?? string.Empty;

            // Extract the section index from the file name (e.g., Section_2.docx -> 2).
            int expectedIndex = int.Parse(Path.GetFileNameWithoutExtension(filePath).Split('_')[1]);

            if (!headerText.Contains($"Header {expectedIndex}") || !footerText.Contains($"Footer {expectedIndex}"))
                throw new InvalidOperationException($"Header/footer verification failed for {filePath}");
        }

        // All split documents have been verified successfully.
        Console.WriteLine("Document splitting and header/footer verification completed successfully.");
    }
}
