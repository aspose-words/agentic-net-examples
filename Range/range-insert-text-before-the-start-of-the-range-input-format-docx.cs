using System;
using Aspose.Words;

class InsertTextBeforeRange
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Get the first node of the document's main story (the start of the range).
        // The Body node of the first section contains the top‑level nodes of the document.
        Node firstNode = doc.FirstSection.Body.FirstChild;

        // Create a Run node that contains the text we want to insert.
        Run newRun = new Run(doc, "Inserted text ");

        // Insert the new Run before the first node of the body.
        // If the body is empty (firstNode == null) we can simply prepend the run.
        if (firstNode != null)
            doc.FirstSection.Body.InsertBefore(newRun, firstNode);
        else
            doc.FirstSection.Body.PrependChild(newRun);

        // Save the modified document.
        doc.Save("output.docx");
    }
}
