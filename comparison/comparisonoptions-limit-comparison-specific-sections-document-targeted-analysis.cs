using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

class CompareSpecificSections
{
    static void Main()
    {
        // Create original document with two sections.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("This is the original first section.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Original second section.");

        // Create edited document with modifications in the first section.
        Document docEdited = new Document();
        DocumentBuilder eb = new DocumentBuilder(docEdited);
        eb.Writeln("This is the edited first section with changes.");
        eb.InsertBreak(BreakType.SectionBreakNewPage);
        eb.Writeln("Original second section."); // unchanged

        // Configure comparison options to focus on main body text only.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = false,
            IgnoreFormatting = true,
            IgnoreCaseChanges = false,
            IgnoreComments = true,
            IgnoreTables = true,
            IgnoreFields = true,
            IgnoreFootnotes = true,
            IgnoreTextboxes = true,
            IgnoreHeadersAndFooters = true,
            Target = ComparisonTargetType.New
        };

        // Extract the first section from each document.
        Document tempOriginal = ExtractSection(docOriginal, 0);
        Document tempEdited   = ExtractSection(docEdited,   0);

        // Perform the comparison.
        tempOriginal.Compare(tempEdited, "Analyzer", DateTime.Now, compareOptions);

        // Save result to the current directory.
        string outPath = Path.Combine(Environment.CurrentDirectory, "ComparisonResult.docx");
        tempOriginal.Save(outPath);
        Console.WriteLine($"Comparison saved to: {outPath}");
    }

    private static Document ExtractSection(Document source, int sectionIndex)
    {
        // Clone the required section.
        Section sectionClone = (Section)source.Sections[sectionIndex].Clone(true);

        // Create a new empty document and import the cloned section.
        Document result = new Document();
        result.RemoveAllChildren();
        Node importedSection = result.ImportNode(sectionClone, true);
        result.AppendChild(importedSection);
        return result;
    }
}
