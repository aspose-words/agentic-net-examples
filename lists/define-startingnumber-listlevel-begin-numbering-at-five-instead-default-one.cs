using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on the default list template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Retrieve the first list level (level 0) and set its starting number to 5.
        ListLevel level = list.ListLevels[0];
        level.StartAt = 5; // List will begin numbering at 5 instead of the default 1.

        // Apply the list to the following paragraphs.
        builder.ListFormat.List = list;

        // Add some list items to demonstrate the custom start number.
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("ListStartAtFive.docx");
    }
}
