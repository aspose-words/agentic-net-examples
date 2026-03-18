using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three separate paragraphs.
        builder.Writeln("First paragraph");
        builder.Writeln("Second paragraph");
        builder.Writeln("Third paragraph");

        // Retrieve all paragraph nodes from the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Create a new list based on a predefined template (bulleted list).
        List sharedList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Apply the same list and the same list level to each paragraph.
        foreach (Paragraph para in paragraphs.OfType<Paragraph>())
        {
            para.ListFormat.List = sharedList;   // Assign the list.
            para.ListFormat.ListLevelNumber = 0; // Use the first level of the list.
        }

        // Save the resulting document.
        doc.Save("Lists_SharedFormatting.docx");
    }
}
