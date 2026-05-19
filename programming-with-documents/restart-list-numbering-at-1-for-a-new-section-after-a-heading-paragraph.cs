using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;

namespace RestartListNumberingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First section – heading and a numbered list.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("First Section");

            // Use the default numbered list template.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // Insert a section break so the next content starts in a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Second section – heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Second Section");

            // Create a list that restarts numbering at each section.
            List restartList = doc.Lists.Add(ListTemplate.NumberDefault);
            restartList.IsRestartAtEachSection = true;

            // Apply the list to the builder and add items.
            builder.ListFormat.List = restartList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.Writeln("Item 3");
            builder.ListFormat.RemoveNumbers();

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "RestartList.docx");
            doc.Save(outputPath);
        }
    }
}
