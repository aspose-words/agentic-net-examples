using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;

namespace ListRestartExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a numbered list based on a built‑in template.
            List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

            // First section – start numbering from 1.
            numberedList.ListLevels[0].StartAt = 1;
            builder.ListFormat.List = numberedList;
            builder.Writeln("Section 1 – Item 1");
            builder.Writeln("Section 1 – Item 2");
            builder.ListFormat.RemoveNumbers();

            // Insert a section break (new page) to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Second section – reset the starting number to 1 again before applying the list.
            numberedList.ListLevels[0].StartAt = 1;
            builder.ListFormat.List = numberedList;
            builder.Writeln("Section 2 – Item 1");
            builder.Writeln("Section 2 – Item 2");
            builder.ListFormat.RemoveNumbers();

            // Save the document to the current directory.
            string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "RestartList.docx");
            doc.Save(outputPath);
        }
    }
}
