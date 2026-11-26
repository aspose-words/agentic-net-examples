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

        // Create a bulleted list using the default bullet template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Assign the created list to the builder.
        builder.ListFormat.List = bulletList;

        // Apply default bullet formatting and write list items.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First bullet item");

        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("Second bullet item");

        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("Third bullet item");

        // Remove list formatting after the list is complete.
        builder.ListFormat.RemoveNumbers();

        // Add a normal paragraph after the list.
        builder.Writeln("This paragraph is not part of the list.");

        // Validate that the document contains the expected number of list paragraphs.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }
        Console.WriteLine($"Number of paragraphs with list formatting: {listParagraphCount}");

        // Save the document to a file.
        doc.Save("BulletedList.docx");
    }
}
