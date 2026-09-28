using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeWordsListCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -------------------------------------------------
            // Create the original list and add some items.
            // -------------------------------------------------
            builder.Writeln("Original List:");
            // Add a numbered list (default template).
            List originalList = doc.Lists.Add(ListTemplate.NumberDefault);
            // Apply the list to subsequent paragraphs.
            builder.ListFormat.List = originalList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.Writeln("Item 3");
            // Stop using the list for further paragraphs.
            builder.ListFormat.RemoveNumbers();

            // -------------------------------------------------
            // Create a new list that will act as a clone.
            // -------------------------------------------------
            List clonedList = doc.Lists.Add(ListTemplate.NumberDefault);

            // Copy formatting from the original list (keeps appearance identical).
            for (int i = 0; i < originalList.ListLevels.Count; i++)
            {
                clonedList.ListLevels[i].NumberStyle = originalList.ListLevels[i].NumberStyle;
                clonedList.ListLevels[i].NumberFormat = originalList.ListLevels[i].NumberFormat;
                clonedList.ListLevels[i].Alignment = originalList.ListLevels[i].Alignment;
                // Font property is read‑only; individual font attributes can be copied if needed.
                // Example: clonedList.ListLevels[i].Font.Name = originalList.ListLevels[i].Font.Name;
            }

            // Modify the start values for each level of the cloned list.
            // For demonstration, set level 0 start at 10, level 1 at 20, etc.
            for (int i = 0; i < clonedList.ListLevels.Count; i++)
            {
                clonedList.ListLevels[i].StartAt = (i + 1) * 10;
            }

            // -------------------------------------------------
            // Add a new section and apply the cloned list there.
            // -------------------------------------------------
            Section newSection = new Section(doc);
            doc.Sections.Add(newSection);

            // Move a builder to the newly added section.
            DocumentBuilder sectionBuilder = new DocumentBuilder(doc);
            sectionBuilder.MoveToSection(doc.Sections.Count - 1);

            sectionBuilder.Writeln("Cloned List with Modified Start Values:");
            // Apply the cloned list to subsequent paragraphs.
            sectionBuilder.ListFormat.List = clonedList;
            sectionBuilder.Writeln("Cloned Item 1");
            sectionBuilder.Writeln("Cloned Item 2");
            sectionBuilder.Writeln("Cloned Item 3");
            sectionBuilder.ListFormat.RemoveNumbers();

            // -------------------------------------------------
            // Save the document to a file.
            // -------------------------------------------------
            doc.Save("ClonedListExample.docx");
        }
    }
}
