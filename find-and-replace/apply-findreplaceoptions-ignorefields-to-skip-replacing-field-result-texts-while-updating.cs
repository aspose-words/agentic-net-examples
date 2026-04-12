using System;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindReplaceIgnoreFieldsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add regular text that will be replaced.
            builder.Writeln("Hello world!");

            // Insert a field whose result also contains the word "Hello".
            // The field code is QUOTE and the result text is "Hello again!".
            builder.InsertField("QUOTE", "Hello again!");

            // Configure find-and-replace options to ignore whole fields.
            // This means the replacement will affect only the regular text,
            // leaving the field result unchanged.
            FindReplaceOptions options = new FindReplaceOptions
            {
                IgnoreFields = true
            };

            // Perform the replacement.
            int replacementCount = doc.Range.Replace("Hello", "Greetings", options);

            // Validate that at least one replacement was made.
            if (replacementCount == 0)
                throw new InvalidOperationException("No replacements were performed.");

            // Save the modified document to the local file system.
            const string outputPath = "Result.docx";
            doc.Save(outputPath);

            // Output a simple confirmation.
            Console.WriteLine($"Replacements made: {replacementCount}");
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
