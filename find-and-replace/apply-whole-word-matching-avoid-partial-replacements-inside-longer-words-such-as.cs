using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class WholeWordReplaceExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample text containing the word "catalogue" and a longer word "catalogue2021".
        builder.Writeln("Our product catalogue is ready.");
        builder.Writeln("The file catalogue2021 will be archived.");

        // Configure find‑replace options to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            FindWholeWordsOnly = true   // Prevents partial matches inside longer words.
        };

        // Replace the whole word "catalogue" with "catalog".
        doc.Range.Replace("catalogue", "catalog", options);

        // Save the result to disk.
        doc.Save("WholeWordReplaceResult.docx");
    }
}
