using System;
using Aspose.Words;
using Aspose.Words.Lists;

class ListExample
{
    static void Main()
    {
        // Path to the existing DOCX file to be loaded.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the modified document will be saved.
        string outputPath = @"C:\Docs\OutputDocument.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new numbered list based on a predefined template.
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to subsequent paragraphs.
        builder.ListFormat.List = numberedList;

        // Set the list level (0 = first level) and add list items.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item 1");
        builder.Writeln("Item 2");
        builder.Writeln("Item 3");

        // Optionally, increase the level to create a sub‑list.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("Sub‑item A");
        builder.Writeln("Sub‑item B");

        // End list formatting for following paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Add a normal paragraph after the list.
        builder.Writeln("This paragraph is not part of the list.");

        // Save the modified document.
        doc.Save(outputPath);
    }
}
