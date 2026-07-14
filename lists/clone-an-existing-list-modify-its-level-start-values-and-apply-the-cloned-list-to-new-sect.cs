using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an original list based on a predefined template.
        List originalList = doc.Lists.Add(ListTemplate.NumberArabicParenthesis);
        // Customize the first level of the original list.
        originalList.ListLevels[0].Font.Color = Color.Red;
        originalList.ListLevels[0].Alignment = ListLevelAlignment.Right;

        // Apply the original list to a few paragraphs.
        builder.Writeln("Original list starts below:");
        builder.ListFormat.List = originalList;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.ListFormat.RemoveNumbers();

        // Clone the existing list using AddCopy (as per Aspose.Words API).
        List clonedList = doc.Lists.AddCopy(originalList);
        // Modify the starting number of the first level in the cloned list.
        clonedList.ListLevels[0].StartAt = 10;
        // Optionally change the color to differentiate the cloned list.
        clonedList.ListLevels[0].Font.Color = Color.Blue;

        // Clone the first section of the document and add it as a new section.
        Section newSection = doc.Sections[0].Clone();
        doc.Sections.Add(newSection);

        // Move the builder cursor to the newly added section.
        DocumentBuilder sectionBuilder = new DocumentBuilder(doc);
        sectionBuilder.MoveToSection(doc.Sections.Count - 1);

        // Apply the cloned list to paragraphs in the new section.
        sectionBuilder.Writeln("Cloned list in a new section starts below:");
        sectionBuilder.ListFormat.List = clonedList;
        sectionBuilder.Writeln("Item 1");
        sectionBuilder.Writeln("Item 2");
        sectionBuilder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("ClonedListExample.docx");
    }
}
