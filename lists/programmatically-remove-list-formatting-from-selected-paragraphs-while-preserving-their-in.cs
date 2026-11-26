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

        // Create a numbered list.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder.
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0; // top level

        // Add three list items.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Remove list formatting from the second paragraph (Item 2) while preserving indentation.
        Paragraph secondParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        DocumentBuilder paraBuilder = new DocumentBuilder(doc);
        paraBuilder.MoveTo(secondParagraph);
        paraBuilder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("output.docx");
    }
}
