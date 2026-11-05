using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class SplitDocumentExample
{
    public static void Main()
    {
        // Create a sample document with sections and heading paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First section with two headings.
        builder.Writeln("Content before first heading in Section 1.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 1.1.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1.2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 1.2.");

        // Start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content before first heading in Section 2.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 2.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Paragraph under Heading 2.1.");

        // Ensure the source document is saved (optional, for inspection).
        sourceDoc.Save("SourceDocument.docx");

        // Directory to hold split results.
        string outputDir = "SplitOutputs";
        Directory.CreateDirectory(outputDir);

        int splitCounter = 1;

        // Split by sections.
        foreach (Section section in sourceDoc.Sections)
        {
            Document sectionDoc = new Document();
            NodeImporter importer = new NodeImporter(sourceDoc, sectionDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(section, true);
            sectionDoc.AppendChild(importedSection);
            string sectionPath = Path.Combine(outputDir, $"Split_{splitCounter++}.docx");
            sectionDoc.Save(sectionPath);
        }

        // Split by headings within each section.
        foreach (Section section in sourceDoc.Sections)
        {
            foreach (Paragraph para in section.Body.Paragraphs)
            {
                if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                {
                    Document headingDoc = new Document();
                    NodeImporter importer = new NodeImporter(sourceDoc, headingDoc, ImportFormatMode.KeepSourceFormatting);

                    // Import the heading paragraph.
                    Paragraph headingClone = (Paragraph)importer.ImportNode(para, true);
                    headingDoc.FirstSection.Body.AppendChild(headingClone);

                    // Import following nodes until the next heading or end of section.
                    Node nextNode = para.NextSibling;
                    while (nextNode != null &&
                           !(nextNode is Paragraph p && p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1))
                    {
                        Node cloned = importer.ImportNode(nextNode, true);
                        headingDoc.FirstSection.Body.AppendChild(cloned);
                        nextNode = nextNode.NextSibling;
                    }

                    string headingPath = Path.Combine(outputDir, $"Split_{splitCounter++}.docx");
                    headingDoc.Save(headingPath);
                }
            }
        }

        // Validation: ensure at least one split file was created.
        string[] splitFiles = Directory.GetFiles(outputDir, "*.docx");
        if (splitFiles.Length == 0)
            throw new InvalidOperationException("No split documents were generated.");

        // Optional: output the count of generated files.
        Console.WriteLine($"Generated {splitFiles.Length} split document(s) in '{outputDir}'.");
    }
}
