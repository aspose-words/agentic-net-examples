using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the formatting for the merge field to bold.
        builder.Font.Bold = true;

        // Insert a merge field named "Name".
        builder.InsertField("MERGEFIELD Name");

        // Prepare the data for the mail merge.
        var fieldNames = new[] { "Name" };
        var fieldValues = new object[] { "John Doe" };

        // Execute the mail merge. The inserted text will inherit the bold formatting.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the resulting document to a file.
        doc.Save("Output.docx");
    }
}
