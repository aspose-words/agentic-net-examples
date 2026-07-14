using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list using the default list template.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the list level to 3 (third‑level item, zero‑based indexing).
        builder.ListFormat.ListLevelNumber = 3;

        // Add a paragraph that will be formatted as a third‑level list item.
        builder.Writeln("Third level list item");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("ThirdLevelList.docx");
    }
}
