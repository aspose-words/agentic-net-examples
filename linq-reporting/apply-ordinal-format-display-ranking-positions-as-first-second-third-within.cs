using System;
using Aspose.Words;
using Aspose.Words.Fields;

class LeaderboardExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title for the leaderboard section.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Leaderboard");

        // Reset to normal style for the list entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Example rankings – you can replace the loop with your actual data source.
        for (int rank = 1; rank <= 5; rank++)
        {
            // Insert a field that formats the numeric rank as ordinal text (First, Second, ...).
            // The field code: QUOTE "rank" \* OrdText
            // QUOTE returns the literal text inside the quotes, and the \* OrdText switch
            // converts the number to its ordinal word representation.
            Field field = builder.InsertField($"QUOTE \"{rank}\" \\* OrdText");

            // Append a description after the formatted rank.
            builder.Writeln($" place – Player {rank}");
        }

        // Save the document to disk.
        doc.Save("Leaderboard.docx");
    }
}
