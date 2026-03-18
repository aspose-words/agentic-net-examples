using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class FindReplaceWholeWordExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample text that contains the target word both as a whole word and as part of another word.
        builder.Writeln("Jackson will meet you in Jacksonville.");
        builder.Writeln("Jackson is a common first name.");

        // Configure find/replace options to match whole words only.
        FindReplaceOptions options = new FindReplaceOptions
        {
            // When true, the old value must be a standalone word.
            FindWholeWordsOnly = true
        };

        // Replace the whole word "Jackson" with "Louis". The occurrence inside "Jacksonville" will be ignored.
        doc.Range.Replace("Jackson", "Louis", options);

        // Save the result to a file (adjust the path as needed).
        doc.Save("FindReplaceWholeWordResult.docx");
    }
}
