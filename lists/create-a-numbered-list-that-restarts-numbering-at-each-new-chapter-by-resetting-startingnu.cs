using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list based on the default template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);
        // Ensure the list restarts at each section (optional, helps when using section breaks).
        list.IsRestartAtEachSection = true;

        // Define three chapters.
        for (int chapter = 1; chapter <= 3; chapter++)
        {
            // Insert a heading for the chapter.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln($"Chapter {chapter}");

            // Start a new section so that the list can restart if needed.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Reset the starting number of the first level to 1 for this chapter.
            list.ListLevels[0].StartAt = 1;

            // Apply the list to subsequent paragraphs.
            builder.ListFormat.List = list;
            builder.ListFormat.ListLevelNumber = 0; // first level

            // Add five items to the list.
            for (int i = 1; i <= 5; i++)
            {
                builder.Writeln($"Item {i} of Chapter {chapter}");
            }

            // End the list for this chapter.
            builder.ListFormat.RemoveNumbers();
        }

        // Save the document.
        doc.Save(Path.Combine(artifactsDir, "NumberedListRestartPerChapter.docx"));
    }
}
