using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a list based on the default numbered template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Access the first list level (level 0).
        ListLevel level = list.ListLevels[0];

        // Set the tab position to 72 points (1 inch) so that the text aligns after the list number.
        level.TabPosition = 72;

        // Ensure the tab character is used as the separator; otherwise TabPosition has no effect.
        level.TrailingCharacter = ListTrailingCharacter.Tab;

        // Use the list in a few paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.ListFormat.List = list;
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("ListWithTabPosition.docx");
    }
}
