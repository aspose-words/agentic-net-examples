using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with sections and headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section with a heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the first chapter.");

        // Add a second heading within the same section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Details of section 1.1.");

        // Insert a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of the second chapter.");

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutput");
        Directory.CreateDirectory(outputDir);

        // ---------- Split by Sections ----------
        int sectionIndex = 1;
        foreach (Section section in doc.Sections)
        {
            Document sectionDoc = new Document();
            // Import the whole section (including headers/footers) into the new document.
            Section importedSection = (Section)sectionDoc.ImportNode(section, true);
            // Remove the default empty section that Document() creates.
            sectionDoc.RemoveAllChildren();
            sectionDoc.AppendChild(importedSection);

            string sectionPath = Path.Combine(outputDir, $"Section_{sectionIndex}.docx");
            sectionDoc.Save(sectionPath);
            sectionIndex++;
        }

        // ---------- Split by Headings ----------
        // Collect all heading paragraphs.
        List<Paragraph> headingParagraphs = new List<Paragraph>();
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
            if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
                headingParagraphs.Add(para);
        }

        for (int i = 0; i < headingParagraphs.Count; i++)
        {
            Paragraph startHeading = headingParagraphs[i];
            Paragraph nextHeading = (i + 1 < headingParagraphs.Count) ? headingParagraphs[i + 1] : null;

            Document headingDoc = new Document();
            NodeImporter importer = new NodeImporter(doc, headingDoc, ImportFormatMode.KeepSourceFormatting);

            // Ensure the document has a section to receive nodes.
            Section bodySection = headingDoc.FirstSection;

            // Import nodes from the start heading up to (but not including) the next heading.
            Node currentNode = startHeading;
            while (currentNode != null && currentNode != nextHeading)
            {
                Node importedNode = importer.ImportNode(currentNode, true);
                bodySection.Body.AppendChild(importedNode);
                currentNode = currentNode.NextSibling;
            }

            string headingPath = Path.Combine(outputDir, $"Heading_{i + 1}.docx");
            headingDoc.Save(headingPath);
        }

        // Validate that split files were created.
        string[] splitFiles = Directory.GetFiles(outputDir, "*.docx");
        if (splitFiles.Length < 2)
            throw new InvalidOperationException($"Expected multiple split files, but found {splitFiles.Length}.");

        Console.WriteLine($"Document split into {splitFiles.Length} files in '{outputDir}'.");
    }
}
