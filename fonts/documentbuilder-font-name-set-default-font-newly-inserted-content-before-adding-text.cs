using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the default font name for all text that will be inserted from this point onward.
        builder.Font.Name = "Arial";

        // Insert text; it will be formatted with the default Arial font.
        builder.Writeln("This paragraph uses the default Arial font.");

        // Change the default font for subsequent text.
        builder.Font.Name = "Times New Roman";

        // Insert more text; this paragraph uses Times New Roman.
        builder.Writeln("This paragraph uses Times New Roman.");

        // Save the document to disk.
        doc.Save("DefaultFontExample.docx");
    }
}
