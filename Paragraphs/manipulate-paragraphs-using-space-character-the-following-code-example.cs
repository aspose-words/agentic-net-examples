using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a three‑level numbered list.
        builder.ListFormat.ApplyNumberDefault(); // start default numbered list
        builder.Writeln("Item 1");               // level 0
        builder.ListFormat.ListIndent();         // increase to level 1
        builder.Writeln("Item 2");               // level 1
        builder.ListFormat.ListIndent();         // increase to level 2
        builder.Writeln("Item 3");               // level 2

        // Configure plain‑text save options to use spaces for list indentation.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        txtOptions.ListIndentation.Character = ' '; // space character for padding
        txtOptions.ListIndentation.Count = 3;      // three spaces per indent level

        // Save the document as a .txt file using the configured options.
        doc.Save("ListIndentation.txt", txtOptions);
    }
}
