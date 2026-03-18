using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class HyperlinkFromDatabase
{
    static void Main()
    {
        // Example URL that the hyperlink will point to.
        string url = "https://www.example.com";

        // Simulate retrieving the display text from a database.
        // In a real scenario replace this with actual DB access logic.
        string dbDisplayText = GetDisplayTextFromDatabase();

        // If the database field is empty or null, fall back to using the URL as the display text.
        string displayText = string.IsNullOrWhiteSpace(dbDisplayText) ? url : dbDisplayText;

        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to insert the hyperlink.
        // The third argument (false) indicates that the second argument is a URL, not a bookmark.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertHyperlink(displayText, url, false);

        // Save the document to disk.
        doc.Save("HyperlinkFromDatabase.docx");
    }

    // Placeholder method that mimics fetching a value from a database.
    // Returns null, empty string, or a meaningful text depending on the test case.
    static string GetDisplayTextFromDatabase()
    {
        // Example: return null; // Uncomment to test fallback to URL.
        // Example: return "";   // Uncomment to test fallback to URL.
        return "Visit Example Site"; // Return a non‑empty value to use as display text.
    }
}
