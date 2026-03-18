using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Example expression that will be part of the URL.
        // In a real scenario this could be built from data, e.g. a product ID.
        string expression = "productId=12345";

        // Build the full external URL using the expression.
        string url = $"https://www.example.com/search?{expression}";

        // Apply hyperlink formatting before inserting the link.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;

        // Insert a hyperlink with display text and the dynamic URL.
        // The third argument (false) indicates that the second argument is a URL, not a bookmark.
        builder.InsertHyperlink("Search for product", url, false);

        // Reset formatting for any following text.
        builder.Font.ClearFormatting();

        // Add a trailing period for readability.
        builder.Writeln(".");

        // Save the document to disk.
        doc.Save("DynamicHyperlink.docx");
    }
}
