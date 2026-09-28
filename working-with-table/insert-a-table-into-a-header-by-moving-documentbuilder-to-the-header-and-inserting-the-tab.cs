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

        // Ensure that the primary header exists.
        // If it does not, create and add it to the first section.
        HeaderFooter primaryHeader = doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
        if (primaryHeader == null)
        {
            primaryHeader = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
            doc.FirstSection.HeadersFooters.Add(primaryHeader);
        }

        // Move the builder's cursor into the primary header.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Build a simple 1‑row, 2‑cell table inside the header.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Write("Header Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.Write("Header Cell 2");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = "HeaderTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");

        // Indicate successful completion.
        Console.WriteLine("Document saved to " + Path.GetFullPath(outputPath));
    }
}
