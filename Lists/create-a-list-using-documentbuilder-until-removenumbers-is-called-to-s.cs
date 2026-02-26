using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document and attach a DocumentBuilder to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a default bulleted list.
        builder.ListFormat.ApplyBulletDefault();

        // Add paragraphs that will become list items.
        builder.Writeln("First bullet item");
        builder.Writeln("Second bullet item");
        builder.Writeln("Third bullet item");

        // End the list – removes bullets and resets the list level to zero.
        builder.ListFormat.RemoveNumbers();

        // Save the resulting DOCX file.
        doc.Save("BulletedList.docx");
    }
}
