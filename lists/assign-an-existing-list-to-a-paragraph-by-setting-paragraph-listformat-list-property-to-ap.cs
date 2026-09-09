using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder for convenient text insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph before the list-assigned paragraph.");

        // Create a list (bulleted by default) that will be assigned to a paragraph later.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Create a paragraph manually (without using the builder) so we can set its ListFormat.
        Paragraph listParagraph = new Paragraph(doc);
        // Add some text to the paragraph.
        Run run = new Run(doc, "This paragraph is part of the existing list.");
        listParagraph.AppendChild(run);

        // Assign the previously created list to the paragraph.
        listParagraph.ListFormat.List = bulletList;
        // Optionally set the list level (0 = first level).
        listParagraph.ListFormat.ListLevelNumber = 0;

        // Append the paragraph to the document body.
        doc.FirstSection.Body.AppendChild(listParagraph);

        // Add another paragraph after the list-assigned one.
        builder.Writeln("Paragraph after the list-assigned paragraph.");

        // Save the document to a file.
        doc.Save("AssignListToParagraph.docx");
    }
}
