using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Add a new list to the document's ListCollection using a bullet list template.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Create first paragraph and assign it to the newly created list.
        Paragraph para1 = new Paragraph(doc);
        para1.AppendChild(new Run(doc, "First paragraph belonging to the shared list."));
        para1.ListFormat.List = list;               // Assign the list.
        para1.ListFormat.ListLevelNumber = 0;       // Use the first level of the list.
        doc.FirstSection.Body.AppendChild(para1);

        // Create second paragraph and assign it to the same list.
        Paragraph para2 = new Paragraph(doc);
        para2.AppendChild(new Run(doc, "Second paragraph belonging to the shared list."));
        para2.ListFormat.List = list;               // Same list as para1.
        para2.ListFormat.ListLevelNumber = 0;
        doc.FirstSection.Body.AppendChild(para2);

        // Save the document to a file.
        doc.Save("ListExample.docx");
    }
}
