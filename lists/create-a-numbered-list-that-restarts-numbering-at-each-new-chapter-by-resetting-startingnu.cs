using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to insert content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list based on the default template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // Helper to start a new chapter.
        void StartChapter(string title)
        {
            // Reset the starting number of the first list level to 1.
            numberedList.ListLevels[0].StartAt = 1;

            // Write the chapter heading (not part of the list).
            builder.ListFormat.RemoveNumbers(); // Ensure heading is not numbered as a list item.
            builder.Writeln(title);

            // Re‑apply the list for the chapter items.
            builder.ListFormat.List = numberedList;
        }

        // Chapter 1
        StartChapter("Chapter 1");
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Insert a section break to separate chapters.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Chapter 2
        StartChapter("Chapter 2");
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");
        builder.Writeln("Item 4");

        // Ensure any remaining list formatting is cleared.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "NumberedListRestartPerChapter.docx");
        doc.Save(outputPath);
    }
}
