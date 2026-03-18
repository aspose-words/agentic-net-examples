using Aspose.Words;
using System;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Remove all characters from the document's range, leaving an empty template.
        doc.Range.Delete();

        // Ensure the document still contains the minimal required nodes for saving.
        doc.EnsureMinimum();

        // Save the empty template to a file.
        doc.Save("EmptyTemplate.docx");
    }
}
