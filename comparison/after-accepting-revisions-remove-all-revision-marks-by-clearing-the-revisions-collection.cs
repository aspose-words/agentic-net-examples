using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("This is the original content.");

        // Create a modified version of the document.
        Document edited = (Document)original.Clone(true);
        // Change the text to introduce a revision.
        edited.FirstSection.Body.FirstParagraph.Runs[0].Text = "This is the edited content.";

        // Compare the documents to generate revisions in the original document.
        // Provide author name and current date/time as required.
        original.Compare(edited, "John Doe", DateTime.Now);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            Console.WriteLine("No revisions were generated; comparison may have failed.");
            return;
        }

        // Accept all revisions, which removes them from the collection.
        original.Revisions.AcceptAll();

        // After acceptance, the revisions collection should be empty.
        if (original.Revisions.Count != 0)
        {
            Console.WriteLine($"Revisions remain after acceptance: {original.Revisions.Count}");
        }

        // Save the final document without any revision marks.
        original.Save("Result.docx");
    }
}
