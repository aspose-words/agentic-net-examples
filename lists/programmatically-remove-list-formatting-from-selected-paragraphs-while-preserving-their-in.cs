using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list and add two items.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;
        builder.Writeln("First item");
        builder.Writeln("Second item");
        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Add a normal paragraph (not part of any list).
        builder.Writeln("A regular paragraph.");

        // Create a bulleted list and add three items.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;
        builder.Writeln("Bullet 1");
        builder.Writeln("Bullet 2");
        builder.Writeln("Bullet 3");
        // End the bulleted list.
        builder.ListFormat.RemoveNumbers();

        // Remove list formatting from the second item of the numbered list
        // while preserving its indentation.
        Paragraph secondNumberedParagraph = doc.FirstSection.Body.Paragraphs[1];
        secondNumberedParagraph.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("ListFormattingRemoved.docx");
    }
}
