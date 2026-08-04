using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample source document with multiple sections.
        //    Each section has its own orientation, header and footer.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1 – Portrait orientation.
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header – Section 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer – Section 1");
        builder.MoveToDocumentStart(); // Return to body.
        builder.Writeln("Content of Section 1 (Portrait).");

        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 – Landscape orientation.
        builder.PageSetup.Orientation = Orientation.Landscape;
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header – Section 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer – Section 2");
        builder.MoveToDocumentStart();
        builder.Writeln("Content of Section 2 (Landscape).");

        // Insert another section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3 – Portrait orientation again.
        builder.PageSetup.Orientation = Orientation.Portrait;
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header – Section 3");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer – Section 3");
        builder.MoveToDocumentStart();
        builder.Writeln("Content of Section 3 (Portrait).");

        // Save the source document (optional, for inspection).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the document by sections and save each part as a DOCX.
        //    Use ImportNode to correctly transfer sections between documents.
        // -----------------------------------------------------------------
        int sectionCount = sourceDoc.Sections.Count;

        for (int i = 0; i < sectionCount; i++)
        {
            // Create a new empty document and remove its default empty section.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren();

            // Import the i‑th section from the source document into the new document.
            // ImportNode performs a deep copy and re‑parents the nodes to the destination document.
            Section importedSection = (Section)partDoc.ImportNode(sourceDoc.Sections[i], true);

            // Append the imported section to the new document.
            partDoc.AppendChild(importedSection);

            // Save the split part.
            string partPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");
            partDoc.Save(partPath);

            // Verify that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split part: {partPath}");
        }

        // -----------------------------------------------------------------
        // 3. Simple confirmation output (no user interaction required).
        // -----------------------------------------------------------------
        Console.WriteLine($"Source document and {sectionCount} split parts have been saved to: {outputDir}");
    }
}
