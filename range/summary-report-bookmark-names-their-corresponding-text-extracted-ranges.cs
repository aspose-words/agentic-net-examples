using System;
using Aspose.Words;
using Aspose.Words.Saving;

class BookmarkSummaryReport
{
    static void Main()
    {
        // Load the source document that contains bookmarks.
        // Replace "Source.docx" with the path to your input file.
        Document sourceDoc = new Document("Source.docx");

        // Get the collection of all bookmarks in the document.
        BookmarkCollection bookmarks = sourceDoc.Range.Bookmarks;

        // Create a new blank document that will hold the summary report.
        Document reportDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(reportDoc);

        // Write a header for the report.
        builder.Writeln("Bookmark Summary Report");
        builder.Writeln("-----------------------");

        // Iterate through each bookmark, extract its name and the text it encloses,
        // and write the information into the report document.
        foreach (Bookmark bookmark in bookmarks)
        {
            // Bookmark.Name gives the bookmark's identifier.
            // Bookmark.Text returns the text enclosed by the bookmark.
            string line = $"Name: {bookmark.Name}\tText: {bookmark.Text}";
            builder.Writeln(line);
        }

        // Save the report document.
        // Replace "BookmarkReport.docx" with the desired output path.
        reportDoc.Save("BookmarkReport.docx", SaveFormat.Docx);

        // Optionally, also output the summary to the console.
        Console.WriteLine("Bookmark summary has been written to BookmarkReport.docx");
    }
}
