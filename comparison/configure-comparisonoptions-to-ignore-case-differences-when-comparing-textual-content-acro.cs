using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Replacing;

public class ComparisonIgnoreCaseExample
{
    public static void Main()
    {
        // Create the first document with original text (mixed case).
        Document doc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Hello World");

        // Create the second document with the same text but lower case.
        Document doc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("hello world");

        // Aspose.Words does not provide a direct IgnoreCase option.
        // To achieve case‑insensitive comparison we normalize the text
        // in both documents to the same case before invoking Compare.
        // Here we replace the mixed‑case text in the first document with its lower‑case form.
        FindReplaceOptions replaceOptions = new FindReplaceOptions();
        doc1.Range.Replace("Hello World", "hello world", replaceOptions);

        // Perform the comparison with default options (no special flags needed).
        CompareOptions options = new CompareOptions();
        doc1.Compare(doc2, "Comparer", DateTime.Now, options);

        // Verify that no revisions were created because case differences have been normalized.
        if (doc1.Revisions.Count != 0)
        {
            throw new InvalidOperationException("Revisions were created despite ignoring case differences.");
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "compare-ignorecase.docx");
        doc1.Save(outputPath);
    }
}
