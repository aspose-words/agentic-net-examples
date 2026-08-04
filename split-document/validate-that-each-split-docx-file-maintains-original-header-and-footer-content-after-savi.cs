using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with two sections, each having its own
        //    header and footer text.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // ----- Section 1 -----
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer 1");

        builder.MoveToSection(0);
        builder.Writeln("Content of section 1.");

        // Insert a section break to start Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ----- Section 2 -----
        // The builder now works on the newly created section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer 2");

        builder.MoveToSection(1);
        builder.Writeln("Content of section 2.");

        // Save the source document.
        string sourcePath = Path.Combine(artifactsDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the document by sections. Each split document will keep
        //    its original header/footer because we copy the whole section.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section originalSection = sourceDoc.Sections[i];

            // Create a new empty document.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren();

            // Import the section into the new document.
            Section importedSection = (Section)splitDoc.ImportNode(originalSection, true, ImportFormatMode.KeepSourceFormatting);
            splitDoc.AppendChild(importedSection);

            // Ensure the document has the minimal structure required.
            splitDoc.EnsureMinimum();

            // Save the split document.
            string splitPath = Path.Combine(artifactsDir, $"Split_{i + 1}.docx");
            splitDoc.Save(splitPath);
        }

        // -----------------------------------------------------------------
        // 3. Validate that each split file exists and that its header/footer
        //    contain the expected text.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string splitPath = Path.Combine(artifactsDir, $"Split_{i + 1}.docx");

            if (!File.Exists(splitPath))
                throw new FileNotFoundException($"Expected split file not found: {splitPath}");

            Document splitDoc = new Document(splitPath);
            HeaderFooter header = splitDoc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            HeaderFooter footer = splitDoc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            string expectedHeader = $"Header {i + 1}";
            string expectedFooter = $"Footer {i + 1}";

            // Header or footer may be null if they were linked to previous; treat missing as failure.
            if (header == null || !header.GetText().Contains(expectedHeader))
                throw new InvalidOperationException($"Header validation failed for {splitPath}. Expected to contain \"{expectedHeader}\".");

            if (footer == null || !footer.GetText().Contains(expectedFooter))
                throw new InvalidOperationException($"Footer validation failed for {splitPath}. Expected to contain \"{expectedFooter}\".");
        }

        // If we reach this point, all validations succeeded.
        Console.WriteLine("All split documents were created and validated successfully.");
    }
}
