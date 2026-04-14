using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some initial content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph without list.");

        // Create a bulleted list using a predefined template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Create a paragraph that will be assigned to the list.
        Paragraph listParagraph = new Paragraph(doc);
        listParagraph.AppendChild(new Run(doc, "This paragraph is part of a bulleted list."));
        // Append the paragraph to the document body.
        doc.FirstSection.Body.AppendChild(listParagraph);

        // Assign the existing list to the paragraph.
        listParagraph.ListFormat.List = bulletList;
        // Set the list level (0 = first level).
        listParagraph.ListFormat.ListLevelNumber = 0;

        // Add another normal paragraph after the list item.
        builder.Writeln("Another normal paragraph.");

        // Save the document.
        string outputPath = "AssignListToParagraph.docx";
        doc.Save(outputPath);
    }
}
