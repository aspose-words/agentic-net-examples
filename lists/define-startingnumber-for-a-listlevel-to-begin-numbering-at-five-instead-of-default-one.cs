using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a new numbered list to the document's list collection.
        // Use a built‑in list template (NumberDefault) as the base.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list to start numbering at five.
        ListLevel level = list.ListLevels[0];
        level.NumberStyle = NumberStyle.Arabic; // Arabic numerals.
        level.StartAt = 5;                     // Start numbering at five.

        // Add first list item.
        Paragraph para1 = new Paragraph(doc);
        para1.ListFormat.List = list;
        para1.ListFormat.ListLevelNumber = 0;
        para1.AppendChild(new Run(doc, "First item"));
        doc.FirstSection.Body.AppendChild(para1);

        // Add second list item.
        Paragraph para2 = new Paragraph(doc);
        para2.ListFormat.List = list;
        para2.ListFormat.ListLevelNumber = 0;
        para2.AppendChild(new Run(doc, "Second item"));
        doc.FirstSection.Body.AppendChild(para2);

        // Save the document to disk.
        doc.Save("ListStartingNumber.docx");
    }
}
