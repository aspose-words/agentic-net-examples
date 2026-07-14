using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few regular paragraphs.
        builder.Writeln("Paragraph without list 1");
        builder.Writeln("Paragraph without list 2");
        builder.Writeln("Paragraph without list 3");

        // Create a bulleted list that we will assign to a paragraph later.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Retrieve the third paragraph (index 2) from the document.
        Paragraph thirdParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 2, true);

        // Assign the existing list to the paragraph and set its level.
        thirdParagraph.ListFormat.List = bulletList;
        thirdParagraph.ListFormat.ListLevelNumber = 0; // first level of the list

        // Continue adding content after the list‑styled paragraph.
        builder.Writeln("Paragraph after list");

        // Save the document to disk.
        doc.Save("AssignListToParagraph.docx");
    }
}
