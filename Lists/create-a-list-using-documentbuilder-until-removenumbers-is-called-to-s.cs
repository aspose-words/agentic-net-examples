using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document – this is the cursor we will use to add content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a default bulleted list. All paragraphs written after this call will become list items.
        builder.ListFormat.ApplyBulletDefault();

        // Write several paragraphs – each becomes a bullet item.
        builder.Writeln("First bullet");
        builder.Writeln("Second bullet");
        builder.Writeln("Third bullet");

        // End the list. This removes the bullet formatting from the current paragraph
        // and resets the list level to zero, so subsequent paragraphs are normal text.
        builder.ListFormat.RemoveNumbers();

        // Save the resulting document to a file.
        doc.Save("BulletedList.docx");
    }
}
