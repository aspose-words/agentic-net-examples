using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with two sections, each having its own header and footer.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1 header/footer and content.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header Text Section 1");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer Text Section 1");
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of Section 1");

        // Insert a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 header/footer and content.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header Text Section 2");
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Footer Text Section 2");
        builder.MoveToDocumentEnd();
        builder.Writeln("Content of Section 2");

        // Save the source document.
        string sourcePath = "Source.docx";
        sourceDoc.Save(sourcePath);

        // Expected header and footer texts for each section.
        string[] expectedHeaders = { "Header Text Section 1", "Header Text Section 2" };
        string[] expectedFooters = { "Footer Text Section 1", "Footer Text Section 2" };

        // Split the document by sections, preserving headers and footers.
        List<string> splitFiles = new List<string>();
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section section = sourceDoc.Sections[i];

            // Create a new document and import the section.
            Document splitDoc = new Document();
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);

            // Replace any existing sections with the imported one.
            splitDoc.Sections.Clear();
            splitDoc.Sections.Add(importedSection);

            string splitPath = $"Section_{i + 1}.docx";
            splitDoc.Save(splitPath);
            splitFiles.Add(splitPath);
        }

        // Validate that each split document contains the original header and footer text.
        for (int i = 0; i < splitFiles.Count; i++)
        {
            string path = splitFiles[i];
            if (!File.Exists(path))
                throw new Exception($"Expected split file not found: {path}");

            Document splitDoc = new Document(path);
            bool headerFound = false;
            bool footerFound = false;

            foreach (HeaderFooter hf in splitDoc.GetChildNodes(NodeType.HeaderFooter, true))
            {
                string text = hf.GetText();
                if (hf.HeaderFooterType == HeaderFooterType.HeaderPrimary && text.Contains(expectedHeaders[i]))
                    headerFound = true;
                if (hf.HeaderFooterType == HeaderFooterType.FooterPrimary && text.Contains(expectedFooters[i]))
                    footerFound = true;
            }

            if (!headerFound)
                throw new Exception($"Header not preserved in {path}");
            if (!footerFound)
                throw new Exception($"Footer not preserved in {path}");
        }

        // All validations passed.
        Console.WriteLine("All split documents preserve their original headers and footers.");
    }
}
