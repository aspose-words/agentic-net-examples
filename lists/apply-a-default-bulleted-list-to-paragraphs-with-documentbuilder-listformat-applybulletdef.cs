using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list using the BulletDefault template.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Assign the list to the builder.
        builder.ListFormat.List = bulletList;

        // First list item.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("First bullet item");

        // Second list item.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("Second bullet item");

        // Third list item.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("Third bullet item");

        // Remove list formatting after the list is complete.
        builder.ListFormat.RemoveNumbers();

        // Validate that the document contains the expected number of list paragraphs.
        int listParagraphCount = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ListFormat.IsListItem)
                listParagraphCount++;
        }

        // Output validation result.
        builder.Writeln($"The document contains {listParagraphCount} bulleted list items.");

        // Save the document to a file.
        doc.Save("BulletedList.docx");
    }
}
