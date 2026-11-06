using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document with two sections.
        //    Each section gets its own header and footer so that they are
        //    preserved when the document is split by sections.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Add header and footer to the first (default) section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header Text");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer Text");

        // Return to the main body and write content for the first section.
        builder.MoveToDocumentEnd();
        builder.Writeln("Section 1 content.");

        // Insert a section break to start a second section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Add the same header and footer to the second section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("Header Text");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("Footer Text");

        // Write content for the second section.
        builder.Writeln("Section 2 content.");

        // Save the source document.
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Split the document by sections, preserving headers and footers.
        // -----------------------------------------------------------------
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new empty document and remove its default empty section.
            Document splitDoc = new Document();
            splitDoc.RemoveAllChildren();

            // Import the current section (including its headers/footers) into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(sourceDoc.Sections[i], true);
            splitDoc.AppendChild(importedSection);

            // Save the split document.
            string splitPath = $"Split_{i + 1}.docx";
            splitDoc.Save(splitPath);

            // -----------------------------------------------------------------
            // 3. Validate that the split document exists and contains the header/footer.
            // -----------------------------------------------------------------
            if (!File.Exists(splitPath))
                throw new Exception($"Split file not found: {splitPath}");

            Document reloaded = new Document(splitPath);
            string headerText = reloaded.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary]?.GetText() ?? "";
            string footerText = reloaded.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary]?.GetText() ?? "";

            if (!headerText.Contains("Header Text"))
                throw new Exception($"Header text missing in {splitPath}");
            if (!footerText.Contains("Footer Text"))
                throw new Exception($"Footer text missing in {splitPath}");
        }

        // All split documents validated successfully.
        Console.WriteLine("All split documents contain the original header and footer.");
    }
}
