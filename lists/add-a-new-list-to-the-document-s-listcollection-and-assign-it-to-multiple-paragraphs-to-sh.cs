using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few regular paragraphs.
        builder.Writeln("Paragraph A");
        builder.Writeln("Paragraph B");
        builder.Writeln("Paragraph C");

        // Create a new list based on a predefined template (bulleted list).
        List sharedList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Assign the same list to all existing paragraphs so they share the same formatting.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs.OfType<Paragraph>())
        {
            para.ListFormat.List = sharedList;
            para.ListFormat.ListLevelNumber = 0; // first level of the list
        }

        // Save the document to disk.
        doc.Save("SharedList.docx");
    }
}
