using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a numbered list based on the default Word template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Optional: customize the appearance of the first list level.
        list.ListLevels[0].Font.Color = Color.DarkBlue;
        list.ListLevels[0].Font.Size = 12;

        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = list;
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document – it can now be printed programmatically or via a dialog
        // using the Aspose.Words reporting engine.
        doc.Save("ReportWithList.docx");
    }
}
