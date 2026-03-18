using System;
using Aspose.Words;
using Aspose.Words.Tables;

class CloneTableExample
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Locate the table to be cloned. Here we take the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Perform a deep clone of the table (including all rows, cells, and contents).
        Table clonedTable = (Table)originalTable.Clone(true);

        // Identify the paragraph after which the cloned table will be inserted.
        // For demonstration we use the second paragraph in the body (index 1).
        Paragraph referenceParagraph = doc.FirstSection.Body.Paragraphs[1];

        // Insert the cloned table immediately after the reference paragraph.
        // InsertAfter is a method of CompositeNode (Body is a CompositeNode).
        doc.FirstSection.Body.InsertAfter(clonedTable, referenceParagraph);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
