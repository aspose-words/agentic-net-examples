using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new blank document and a builder to work with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a default numbered list (level 0 uses Arabic numbers).
        builder.ListFormat.ApplyNumberDefault();

        // Configure the second list level (index 1) to use lowercase letters
        // and increase its indentation.
        // The ListLevel property returns the formatting of the current level.
        builder.ListFormat.ListLevel.NumberStyle = NumberStyle.LowercaseLetter; // a., b., c., ...
        builder.ListFormat.ListLevel.NumberPosition = 72;   // position of the label (points)
        builder.ListFormat.ListLevel.TextPosition = 108;   // start of the text after the label

        // First‑level items (Arabic numbers).
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");

        // Increase the list level to the second level we just configured.
        builder.ListFormat.ListIndent();

        // Second‑level items (lowercase letters).
        builder.Writeln("Subitem a");
        builder.Writeln("Subitem b");
        builder.Writeln("Subitem c");

        // Return to the first level.
        builder.ListFormat.ListOutdent();

        // Continue first‑level items.
        builder.Writeln("Item 3");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("SecondLevelLowercase.docx");
    }
}
