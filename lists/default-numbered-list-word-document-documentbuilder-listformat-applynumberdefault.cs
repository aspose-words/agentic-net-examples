using System;
using Aspose.Words;
using Aspose.Words.Lists;

class CreateNumberedList
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Add list items. Each Writeln call creates a new paragraph that is part of the list.
        builder.Writeln("First item");
        builder.Writeln("Second item");
        builder.Writeln("Third item");

        // End the numbered list.
        builder.ListFormat.RemoveNumbers();

        // Save the document to a file.
        doc.Save("NumberedList.docx");
    }
}
