using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Load an existing DOCX file
        string inputPath = @"C:\Docs\Input.docx";
        string outputPath = @"C:\Docs\Output.docx";

        Document doc = new Document(inputPath);
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new numbered list based on the default template
        List numberedList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to subsequent paragraphs
        builder.ListFormat.List = numberedList;

        // Add list items
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Item {i}");
        }

        // End the list formatting
        builder.ListFormat.RemoveNumbers();

        // Save the modified document
        doc.Save(outputPath);
    }
}
