using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add two plain paragraphs that initially have no list formatting.
        builder.Writeln("First paragraph without list.");
        builder.Writeln("Second paragraph without list.");

        // Create a bulleted list that will be applied to existing paragraphs.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

        // Retrieve the first paragraph from the document.
        Paragraph firstParagraph = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[0];
        // Assign the list to the paragraph and set its list level (0 = top level).
        firstParagraph.ListFormat.List = bulletList;
        firstParagraph.ListFormat.ListLevelNumber = 0;

        // Retrieve the second paragraph.
        Paragraph secondParagraph = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];
        // Assign the same list but use a deeper level (1 = second level).
        secondParagraph.ListFormat.List = bulletList;
        secondParagraph.ListFormat.ListLevelNumber = 1;

        // Save the resulting document.
        doc.Save("AssignListToParagraph.docx");
    }
}
