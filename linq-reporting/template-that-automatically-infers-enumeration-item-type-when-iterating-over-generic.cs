using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;

public static class ListDocumentGenerator
{
    // Generates a Word document that contains each item of a generic collection as a separate paragraph.
    // The type of the items is inferred automatically by the compiler using 'var' in the foreach loop.
    public static Document GenerateDocumentFromList<T>(IEnumerable<T> items)
    {
        // Create a new empty document.
        Document doc = new Document();

        // Obtain a DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optionally, start a bulleted list to format the items.
        builder.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.ListLevelNumber = 0;

        // Iterate over the generic collection. 'var' lets the compiler infer the item type.
        foreach (var item in items)
        {
            // Write the string representation of each item as a list entry.
            builder.Writeln(item?.ToString() ?? string.Empty);
        }

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        return doc;
    }

    // Example usage.
    public static void Main()
    {
        // Sample generic list of different types.
        List<int> numbers = new List<int> { 1, 2, 3, 4, 5 };
        List<string> words = new List<string> { "Apple", "Banana", "Cherry" };
        List<DateTime> dates = new List<DateTime>
        {
            DateTime.Now,
            DateTime.UtcNow,
            new DateTime(2020, 1, 1)
        };

        // Generate documents for each list.
        Document numbersDoc = GenerateDocumentFromList(numbers);
        Document wordsDoc = GenerateDocumentFromList(words);
        Document datesDoc = GenerateDocumentFromList(dates);

        // Save the documents (paths can be adjusted as needed).
        numbersDoc.Save("NumbersList.docx");
        wordsDoc.Save("WordsList.docx");
        datesDoc.Save("DatesList.docx");
    }
}
