using System;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.BuildingBlocks;

class IncludeHeaderExample
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a reusable header fragment document.
        // -----------------------------------------------------------------
        Document headerFragment = new Document();
        DocumentBuilder fragBuilder = new DocumentBuilder(headerFragment);

        // Add common header content.
        fragBuilder.Writeln("Company Name");
        fragBuilder.Writeln("1234 Main Street");
        fragBuilder.Writeln("City, State ZIP");

        // Save the fragment to a file that will be referenced later.
        const string headerFragmentPath = "HeaderFragment.docx";
        headerFragment.Save(headerFragmentPath);

        // -----------------------------------------------------------------
        // 2. Create the main document and insert the INCLUDE field.
        // -----------------------------------------------------------------
        Document mainDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Move the cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Insert an INCLUDE field that pulls in the header fragment.
        FieldInclude includeField = (FieldInclude)builder.InsertField(FieldType.FieldInclude, true);
        includeField.SourceFullName = headerFragmentPath;   // Path to the fragment.
        includeField.LockFields = false;                    // Allow fields inside the fragment to update.
        includeField.TextConverter = "Microsoft Word";      // Use Word's text converter.

        // Return to the main body and add some sample content.
        builder.MoveToDocumentEnd();
        builder.Writeln("This is the main document body.");

        // Update all fields so the INCLUDE field resolves.
        mainDoc.UpdateFields();

        // Save the final document.
        const string resultPath = "Result.docx";
        mainDoc.Save(resultPath);
    }
}
