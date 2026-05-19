using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for output split documents.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with two sections, each having
        //    its own header and footer.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Content of Section 1");
        AddHeaderFooter(sourceDoc, "Header 1", "Footer 1");

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Content of Section 2");
        AddHeaderFooter(sourceDoc, "Header 2", "Footer 2");

        // Save the source document (optional, for inspection).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the document by sections, preserving headers/footers.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the current section from the source document into the new document.
            // ImportNode clones the node and reassigns it to the target document,
            // which avoids the cross‑document node exception.
            Section importedSection = (Section)splitDoc.ImportNode(sourceDoc.Sections[i], true);

            // Append the imported section as the sole section of the split document.
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");
            splitDoc.Save(splitPath);
        }

        // -----------------------------------------------------------------
        // 3. Verify that each split document preserves its header and footer.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string splitPath = Path.Combine(outputDir, $"Section_{i + 1}.docx");
            if (!File.Exists(splitPath))
                throw new FileNotFoundException($"Expected split file not found: {splitPath}");

            Document part = new Document(splitPath);

            // Retrieve header and footer text (Primary type is used in this example).
            string headerText = part.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]?.GetText()?.Trim() ?? string.Empty;
            string footerText = part.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary]?.GetText()?.Trim() ?? string.Empty;

            string expectedHeader = $"Header {i + 1}";
            string expectedFooter = $"Footer {i + 1}";

            if (!headerText.Contains(expectedHeader))
                throw new Exception($"Header verification failed for {splitPath}. Expected to contain \"{expectedHeader}\", but got \"{headerText}\".");

            if (!footerText.Contains(expectedFooter))
                throw new Exception($"Footer verification failed for {splitPath}. Expected to contain \"{expectedFooter}\", but got \"{footerText}\".");
        }

        // If we reach this point, all verifications succeeded.
        Console.WriteLine("All split documents were created and verified successfully.");
    }

    // Helper method to add a primary header and footer to the most recent section.
    private static void AddHeaderFooter(Document doc, string headerContent, string footerContent)
    {
        // The most recent section is the last one in the collection.
        Section currentSection = doc.Sections[doc.Sections.Count - 1];

        // Header
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
        header.AppendParagraph(headerContent);
        currentSection.HeadersFooters.Add(header);

        // Footer
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
        footer.AppendParagraph(footerContent);
        currentSection.HeadersFooters.Add(footer);
    }
}
