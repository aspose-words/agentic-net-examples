using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder for convenience when adding non‑list content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph before the list.");

        // Create a list based on a predefined template (bulleted list).
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // First list item: create a paragraph, add a run, and assign the list.
        Paragraph para1 = new Paragraph(doc);
        para1.AppendChild(new Run(doc, "First list item"));
        para1.ListFormat.List = list;               // Assign the existing list.
        para1.ListFormat.ListLevelNumber = 0;       // Use the first level of the list.
        doc.FirstSection.Body.AppendChild(para1);

        // Second list item: same steps as above.
        Paragraph para2 = new Paragraph(doc);
        para2.AppendChild(new Run(doc, "Second list item"));
        para2.ListFormat.List = list;
        para2.ListFormat.ListLevelNumber = 0;
        doc.FirstSection.Body.AppendChild(para2);

        // Add a normal paragraph after the list.
        builder.Writeln("Paragraph after the list.");

        // Save the document.
        doc.Save("AssignListToParagraph.docx");
    }
}
