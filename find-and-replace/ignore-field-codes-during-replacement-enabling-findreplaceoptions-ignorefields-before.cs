using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field whose field code contains the letter 'T'.
        // The field result will be "Test IT!".
        builder.InsertField("INCLUDETEXT", "Test IT!");

        // Set up find/replace options to ignore whole fields.
        FindReplaceOptions options = new FindReplaceOptions
        {
            IgnoreFields = true   // Do not search inside field start/end nodes.
        };

        // Replace every occurrence of the character 'T' with '*'.
        // Because IgnoreFields is true, the 'T' inside the field will be left unchanged.
        doc.Range.Replace("T", "*", options);

        // Write the resulting document text to the console.
        Console.WriteLine(doc.GetText());
    }
}
