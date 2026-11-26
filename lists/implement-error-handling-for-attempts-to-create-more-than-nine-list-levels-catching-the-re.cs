using System;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list using a built‑in template.
        List list = doc.Lists.Add(ListTemplate.NumberDefault);

        // Apply the list to the builder.
        builder.ListFormat.List = list;

        // Write a normal list item at level 0.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("Item level 0");

        // Attempt to set an invalid list level (greater than the maximum of 8).
        try
        {
            builder.ListFormat.ListLevelNumber = 9; // Invalid level.
            builder.Writeln("This text will not be added.");
        }
        catch (ArgumentOutOfRangeException ex)
        {
            // Handle the exception thrown by Aspose.Words.
            Console.WriteLine("Caught exception: " + ex.Message);
        }

        // Remove list formatting after the list is complete.
        builder.ListFormat.RemoveNumbers();

        // Save the document.
        doc.Save("output.docx");
    }
}
