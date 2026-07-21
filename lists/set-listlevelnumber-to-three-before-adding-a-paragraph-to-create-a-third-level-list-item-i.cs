using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a numbered list using the default list template.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.NumberDefault);

        // Set the list level to the third level (zero‑based index 2).
        builder.ListFormat.ListLevelNumber = 2;

        // Add a paragraph that will be formatted as a third‑level list item.
        builder.Writeln("Third level list item");

        // Remove list formatting from subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the document to the current directory.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "ThirdLevelList.docx");
        doc.Save(outputPath);
    }
}
