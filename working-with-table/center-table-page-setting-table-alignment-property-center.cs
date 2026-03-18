using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing document that contains at least one table.
        Document doc = new Document("Input.docx");

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Center the table on the page.
        table.Alignment = TableAlignment.Center;

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
