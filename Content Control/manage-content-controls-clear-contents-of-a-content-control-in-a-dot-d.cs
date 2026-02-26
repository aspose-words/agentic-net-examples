using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the DOT template document.
        // This uses the required load rule (Document(string)).
        Document doc = new Document("Template.dot");

        // Iterate over all content controls (structured document tags) in the document.
        // The Clear() method removes the contents and shows the placeholder if defined.
        foreach (StructuredDocumentTag tag in doc.Range.StructuredDocumentTags)
        {
            tag.Clear();
        }

        // Save the modified document.
        // This uses the required save rule (Document.Save(string)).
        doc.Save("Result.docx");
    }
}
