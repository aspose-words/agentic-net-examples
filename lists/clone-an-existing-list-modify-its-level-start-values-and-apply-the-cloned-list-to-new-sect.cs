using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -----------------------------------------------------------------
            // Create the original list and add a couple of items.
            // -----------------------------------------------------------------
            List originalList = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
            originalList.ListLevels[0].Font.Color = Color.Red; // Red numbers for the original list.

            builder.Writeln("Original List:");
            builder.ListFormat.List = originalList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // -----------------------------------------------------------------
            // Clone the original list, modify its start value and color.
            // -----------------------------------------------------------------
            List clonedList = doc.Lists.AddCopy(originalList);
            clonedList.ListLevels[0].StartAt = 10;          // Restart numbering at 10.
            clonedList.ListLevels[0].Font.Color = Color.Blue; // Blue numbers for the cloned list.

            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Apply the cloned list in the new section.
            builder.Writeln("Cloned List (starts at 10):");
            builder.ListFormat.List = clonedList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // -----------------------------------------------------------------
            // Demonstrate cloning an entire section.
            // -----------------------------------------------------------------
            // Clone the first section (which contains the original list) and add it to the end of the document.
            Section firstSection = doc.Sections[0];
            Section duplicatedSection = firstSection.Clone();
            doc.Sections.Add(duplicatedSection);

            // Save the document.
            doc.Save("ClonedListExample.docx");
        }
    }
}
