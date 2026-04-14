using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

namespace ListCloneExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -----------------------------------------------------------------
            // 1. Create an original list and add a few items.
            // -----------------------------------------------------------------
            List originalList = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
            // Customize the first level of the original list.
            originalList.ListLevels[0].Font.Color = Color.Red;
            originalList.ListLevels[0].Alignment = ListLevelAlignment.Right;

            // Apply the original list to some paragraphs.
            builder.Writeln("Original list starts below:");
            builder.ListFormat.List = originalList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // -----------------------------------------------------------------
            // 2. Clone the existing list, modify its start value, and use it.
            // -----------------------------------------------------------------
            List clonedList = doc.Lists.AddCopy(originalList);
            // Change the start number for the first level of the cloned list.
            clonedList.ListLevels[0].StartAt = 10;
            // Optionally change the color to differentiate.
            clonedList.ListLevels[0].Font.Color = Color.Blue;

            // Insert a section break to start a new section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Apply the cloned list to new paragraphs in the new section.
            builder.Writeln("Cloned list (starting at 10) starts below:");
            builder.ListFormat.List = clonedList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // -----------------------------------------------------------------
            // 3. Clone the first section and append it to the document.
            // -----------------------------------------------------------------
            // The first section (index 0) contains the original list items.
            Section firstSection = doc.Sections[0];
            Section duplicatedSection = firstSection.Clone();
            doc.Sections.Add(duplicatedSection);

            // Save the document to a file.
            doc.Save("ListCloneExample.docx");
        }
    }
}
