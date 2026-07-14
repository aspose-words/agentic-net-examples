using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(original);
        builder1.Writeln("Hello world.");

        // Create the revised document with a deliberate change.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);
        builder2.Writeln("Hello revised world.");

        // Compare the documents. Revisions are added to the original document.
        original.Compare(revised, "Author", DateTime.Now);

        // Ensure that at least one revision was generated.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the result, which includes the tracked revisions.
        original.Save("Compared.docx");
    }
}
