using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;

namespace RestartListNumberingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a numbered list and configure it to restart at each new section.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);
            list.IsRestartAtEachSection = true;
            builder.ListFormat.List = list;

            // First section heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Section 1");

            // List items for the first section.
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            // Insert a section break (new page) to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Second section heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Section 2");

            // List items for the second section – numbering restarts at 1.
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            // Save the document to the local file system.
            string outputPath = "RestartListNumbering.docx";
            doc.Save(outputPath);
        }
    }
}
