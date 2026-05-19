using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Section 1 ----------
        builder.Writeln("Section 1 - Original List:");

        // Create a list based on a predefined template.
        List originalList = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
        // Apply the list to a couple of paragraphs.
        builder.ListFormat.List = originalList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // Insert a section break to start a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // ---------- Clone the first section ----------
        // Clone the first section (which contains the original list).
        Section clonedSection = doc.Sections[0].Clone();
        // Add the cloned section to the end of the document.
        doc.Sections.Add(clonedSection);

        // Create a copy of the original list.
        List clonedList = doc.Lists.AddCopy(originalList);
        // Modify the starting number of the first level.
        clonedList.ListLevels[0].StartAt = 10;
        // Optionally change the appearance of the cloned list.
        clonedList.ListLevels[0].Font.Color = Color.Blue;

        // Apply the cloned list to the list items inside the cloned section.
        foreach (Paragraph paragraph in clonedSection.Body.Paragraphs)
        {
            if (paragraph.ListFormat.IsListItem)
            {
                // Replace the list reference with the cloned list.
                paragraph.ListFormat.List = clonedList;
                // Preserve the original list level number.
                paragraph.ListFormat.ListLevelNumber = paragraph.ListFormat.ListLevelNumber;
            }
        }

        // Save the document to disk.
        doc.Save("ClonedListExample.docx");
    }
}
