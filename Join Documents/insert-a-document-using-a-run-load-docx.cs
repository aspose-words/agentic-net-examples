using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX file whose content we want to insert.
        Document srcDoc = new Document("Source.docx");

        // Extract the plain text from the source document.
        string srcText = srcDoc.GetText();

        // Create a new blank document that will receive the inserted content.
        Document dstDoc = new Document();

        // Ensure the document has at least one paragraph to host the run.
        dstDoc.EnsureMinimum();

        // Create a Run that belongs to the destination document and contains the source text.
        Run run = new Run(dstDoc, srcText);

        // Append the run to the first paragraph of the destination document.
        dstDoc.FirstSection.Body.FirstParagraph.AppendChild(run);

        // Save the resulting document.
        dstDoc.Save("Result.docx");
    }
}
