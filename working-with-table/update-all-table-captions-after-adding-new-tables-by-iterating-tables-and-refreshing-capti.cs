using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Helper to add a caption paragraph followed by a simple 2x2 table.
        void AddTableWithCaption(string captionText)
        {
            builder.Writeln(captionText);
            builder.StartTable();
            for (int r = 0; r < 2; r++)
            {
                for (int c = 0; c < 2; c++)
                {
                    builder.InsertCell();
                    builder.Writeln($"R{r + 1}C{c + 1}");
                }
                builder.EndRow();
            }
            builder.EndTable();
        }

        // Add three initial tables with captions.
        for (int i = 1; i <= 3; i++)
        {
            AddTableWithCaption($"Table {i}: Sample caption");
        }

        // Insert a new table at the beginning of the document.
        builder.MoveToDocumentStart();
        AddTableWithCaption("Table X: New table");

        // Refresh all table captions by iterating tables in document order.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = (Table)tables[i];

            // Find the preceding paragraph that holds the caption.
            Node prevNode = table.PreviousSibling;
            while (prevNode != null && !(prevNode is Paragraph))
                prevNode = prevNode.PreviousSibling;

            if (prevNode is Paragraph captionParagraph)
            {
                // Replace the caption text with the correct sequential number.
                captionParagraph.Runs.Clear();
                captionParagraph.AppendChild(new Run(doc, $"Table {i + 1}: Updated caption"));
            }
        }

        // Save the document.
        string outputPath = "UpdatedCaptions.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // Confirmation message.
        Console.WriteLine($"Document saved to '{Path.GetFullPath(outputPath)}' with refreshed table captions.");
    }
}
