using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class MailMergeTemplateCreator
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a static header to the document.
        builder.Writeln("=== Company Confidential ===");
        builder.Writeln(); // Empty line for spacing.

        // Insert a title for the mail merge document.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Customer Information");

        // Reset style for the body.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Insert merge fields.
        builder.Writeln("First Name: ");
        builder.InsertField("MERGEFIELD FirstName \\* MERGEFORMAT");
        builder.Writeln();

        builder.Writeln("Last Name: ");
        builder.InsertField("MERGEFIELD LastName \\* MERGEFORMAT");
        builder.Writeln();

        builder.Writeln("Email: ");
        builder.InsertField("MERGEFIELD Email \\* MERGEFORMAT");
        builder.Writeln();

        // Save the template to a file.
        string outputPath = "MailMergeTemplate.docx";
        doc.Save(outputPath);
    }
}
