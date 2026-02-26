using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToDotm
{
    static void Main()
    {
        // Path to the source DOTM template.
        string dataDir = @"C:\Docs\";
        string inputPath = dataDir + "Template.dotm";

        // Load the DOTM document.
        Document doc = new Document(inputPath);

        // Ensure the document contains at least one table.
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            // Get the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // Create a new row belonging to the same document.
            Row newRow = new Row(doc);

            // Ensure the row has at least one cell.
            newRow.EnsureMinimum();

            // Add content to the first cell of the new row.
            // Here we insert a paragraph with some sample text.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, "New row added programmatically.");
            para.AppendChild(run);
            newRow.FirstCell.FirstParagraph.AppendChild(para);

            // Append the new row to the end of the table.
            table.AppendChild(newRow);
        }
        else
        {
            Console.WriteLine("The document does not contain any tables.");
        }

        // Save the modified document as a DOTM file.
        string outputPath = dataDir + "Modified.dotm";
        doc.Save(outputPath, SaveFormat.Dotm);
    }
}
