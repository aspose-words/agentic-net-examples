using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a custom list based on a built‑in template.
        List customList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first level of the list.
        ListLevel level0 = customList.ListLevels[0];
        level0.NumberFormat = "%1.";
        level0.Alignment = ListLevelAlignment.Left;
        level0.NumberStyle = NumberStyle.Arabic;
        level0.NumberPosition = 0;
        level0.TextPosition = 30;

        // Configure the second level of the list.
        ListLevel level1 = customList.ListLevels[1];
        level1.NumberFormat = "%1.%2.";
        level1.Alignment = ListLevelAlignment.Left;
        level1.NumberStyle = NumberStyle.Arabic;
        level1.NumberPosition = 30;
        level1.TextPosition = 60;

        // Add a paragraph that uses the first level of the list.
        Paragraph para1 = new Paragraph(doc);
        para1.ListFormat.List = customList;
        para1.ListFormat.ListLevelNumber = 0;
        para1.AppendChild(new Run(doc, "First item"));
        doc.FirstSection.Body.AppendChild(para1);

        // Add a paragraph that uses the second level of the list.
        Paragraph para2 = new Paragraph(doc);
        para2.ListFormat.List = customList;
        para2.ListFormat.ListLevelNumber = 1;
        para2.AppendChild(new Run(doc, "Second level item"));
        doc.FirstSection.Body.AppendChild(para2);

        // Save the document to a file.
        doc.Save("CustomList.docx");
    }
}
