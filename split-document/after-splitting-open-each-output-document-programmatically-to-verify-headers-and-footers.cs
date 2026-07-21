using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Remove any previous split files.
        foreach (string file in Directory.GetFiles(outputDir, "*.docx"))
            File.Delete(file);

        // -------------------------------------------------------------
        // 1. Create a sample source document with three sections.
        //    Each section has distinct header/footer texts.
        // -------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Enable different headers/footers for first, even and odd pages.
        builder.PageSetup.DifferentFirstPageHeaderFooter = true;
        builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

        for (int i = 1; i <= 3; i++)
        {
            // Header – Primary
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln($"Header Primary Sec{i}");

            // Header – Even
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Writeln($"Header Even Sec{i}");

            // Header – First
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Writeln($"Header First Sec{i}");

            // Footer – Primary
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln($"Footer Primary Sec{i}");

            // Footer – Even
            builder.MoveToHeaderFooter(HeaderFooterType.FooterEven);
            builder.Writeln($"Footer Even Sec{i}");

            // Footer – First
            builder.MoveToHeaderFooter(HeaderFooterType.FooterFirst);
            builder.Writeln($"Footer First Sec{i}");

            // Return to the main body and add some content.
            builder.MoveToDocumentEnd();
            builder.Writeln($"Content of section {i}");

            // Insert a section break after each section except the last.
            if (i < 3)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Save the source document (optional, for reference).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------------------
        // 2. Split the document by sections, preserving headers/footers.
        // -------------------------------------------------------------
        for (int idx = 0; idx < sourceDoc.Sections.Count; idx++)
        {
            // Create a new empty document.
            Document splitDoc = new Document();
            // Remove the automatically created empty section.
            splitDoc.Sections.Clear();

            // Import the required section from the source document into the new document.
            // ImportNode clones the node and reassigns it to the target document.
            Section importedSection = (Section)splitDoc.ImportNode(sourceDoc.Sections[idx], true);
            splitDoc.Sections.Add(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(outputDir, $"Section_{idx + 1}.docx");
            splitDoc.Save(splitPath);
        }

        // -------------------------------------------------------------
        // 3. Verify that each split document contains the expected headers
        //    and footers.
        // -------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            string splitPath = Path.Combine(outputDir, $"Section_{i}.docx");
            if (!File.Exists(splitPath))
                throw new FileNotFoundException($"Expected split file not found: {splitPath}");

            Document part = new Document(splitPath);
            Section sec = part.FirstSection;

            // Helper to check that a header/footer contains the expected marker text.
            void Verify(HeaderFooterType type, string expected)
            {
                HeaderFooter hf = sec.HeadersFooters[type];
                if (hf == null || !hf.GetText().Contains(expected))
                    throw new InvalidOperationException(
                        $"Header/Footer of type {type} does not contain expected text \"{expected}\" in {splitPath}");
            }

            Verify(HeaderFooterType.HeaderPrimary, $"Header Primary Sec{i}");
            Verify(HeaderFooterType.HeaderEven, $"Header Even Sec{i}");
            Verify(HeaderFooterType.HeaderFirst, $"Header First Sec{i}");
            Verify(HeaderFooterType.FooterPrimary, $"Footer Primary Sec{i}");
            Verify(HeaderFooterType.FooterEven, $"Footer Even Sec{i}");
            Verify(HeaderFooterType.FooterFirst, $"Footer First Sec{i}");
        }

        Console.WriteLine("Document split and verification completed successfully.");
    }
}
