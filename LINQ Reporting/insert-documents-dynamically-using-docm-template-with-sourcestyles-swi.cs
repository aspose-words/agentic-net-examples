using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class DynamicDocumentInserter
{
    static void Main()
    {
        // Path to the DOCM template that contains a bookmark named "InsertHere".
        string templatePath = @"C:\Docs\Template.docm";

        // Load the DOCM template.
        Document templateDoc = new Document(templatePath);

        // Create a DocumentBuilder for the template.
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // List of source documents to insert dynamically.
        List<string> sourcePaths = new List<string>
        {
            @"C:\Docs\Part1.docx",
            @"C:\Docs\Part2.docx",
            @"C:\Docs\Part3.docx"
        };

        // Prepare import options with SmartStyleBehavior enabled.
        // This makes the insertion respect source styles while avoiding style name clashes.
        ImportFormatOptions importOptions = new ImportFormatOptions
        {
            SmartStyleBehavior = true
        };

        // Insert each source document at the bookmark location.
        foreach (string srcPath in sourcePaths)
        {
            // Load the source document.
            Document srcDoc = new Document(srcPath);

            // Move the cursor to the bookmark where the insertion should occur.
            builder.MoveToBookmark("InsertHere");

            // Insert the source document using KeepSourceFormatting to preserve its formatting.
            // SmartStyleBehavior will convert conflicting styles to direct formatting.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting, importOptions);
        }

        // Optionally, attach the original template to the resulting document
        // and enable automatic style updates when the document is opened in Word.
        templateDoc.AttachedTemplate = templatePath;
        templateDoc.AutomaticallyUpdateStyles = true;

        // Save the combined document.
        string outputPath = @"C:\Docs\CombinedResult.docx";
        templateDoc.Save(outputPath);
    }
}
