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
        builder.Writeln("Paragraph before the list.");

        // Create a list (bulleted default) and keep a reference to it.
        List list = doc.Lists.Add(ListTemplate.BulletDefault);

        // Create a new paragraph that will be part of the list.
        Paragraph listParagraph = new Paragraph(doc);
        listParagraph.AppendChild(new Run(doc, "This paragraph is a list item."));

        // Assign the existing list to the paragraph.
        listParagraph.ListFormat.List = list;
        // Set the list level (0 = first level).
        listParagraph.ListFormat.ListLevelNumber = 0;

        // Append the list paragraph to the document body.
        doc.FirstSection.Body.AppendChild(listParagraph);

        // Add another normal paragraph after the list.
        builder.Writeln("Paragraph after the list.");

        // Save the document to a file.
        doc.Save("AssignListToParagraph.docx");
    }
}
