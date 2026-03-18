using System;
using Aspose.Words;
using Aspose.Words.Lists;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a list based on the default numbered template.
        List rankingList = doc.Lists.Add(ListTemplate.NumberDefault);

        // Configure the first list level to display ordinal text (First, Second, Third, …).
        ListLevel level = rankingList.ListLevels[0];
        level.NumberStyle = NumberStyle.OrdinalText;   // Use ordinal text format.
        level.NumberFormat = "\x0000";                // Placeholder for the generated number.
        level.StartAt = 1;                            // Start counting from 1.

        // Apply the list to the builder so subsequent paragraphs become list items.
        builder.ListFormat.List = rankingList;

        // Insert ranking entries. The list will automatically render "First", "Second", "Third", etc.
        builder.Writeln("Alice – 150 points");
        builder.Writeln("Bob – 120 points");
        builder.Writeln("Charlie – 100 points");

        // Stop list formatting for any following content.
        builder.ListFormat.RemoveNumbers();

        // Save the document to disk.
        doc.Save("Ranking.docx");
    }
}
