using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

class IncludeHtmlDynamic
{
    static void Main()
    {
        // The category that determines which HTML fragment to include.
        string category = "CategoryA";

        // Build the full path to the external HTML file based on the category.
        // Example: "C:\Data\Fragments\CategoryA.html"
        string fragmentsDir = @"C:\Data\Fragments";
        string htmlFilePath = Path.Combine(fragmentsDir, $"{category}.html");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an INCLUDETEXT field that will pull in the HTML fragment.
        // The field type for INCLUDETEXT is FieldIncludeText.
        FieldIncludeText includeField = (FieldIncludeText)builder.InsertField(FieldType.FieldIncludeText, true);

        // Set the source file to the HTML fragment we built above.
        includeField.SourceFullName = htmlFilePath;

        // Tell Word to treat the included file as HTML.
        // The TextConverter property defines how the source file is interpreted.
        includeField.TextConverter = "HTML";

        // Optional: prevent fields inside the included fragment from being updated.
        includeField.LockFields = false;

        // Update all fields so the INCLUDETEXT result is calculated.
        doc.UpdateFields();

        // Save the resulting document.
        string outputPath = @"C:\Output\Result.docx";
        doc.Save(outputPath);
    }
}
