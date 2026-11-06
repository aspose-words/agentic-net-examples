using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple template with merge fields.
        builder.Writeln("Hello ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Writeln(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln("!");

        // Prepare data for the merge.
        string[] fieldNames = { "FirstName", "LastName" };
        object[] fieldValues = { "John", "Doe" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document as DOCX.
        string outputPath = "MergedDocument.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The merged document was not created.");

        // Load the saved document and verify that merge fields were replaced.
        Document resultDoc = new Document(outputPath);
        string resultText = resultDoc.GetText();

        if (!resultText.Contains("John") || !resultText.Contains("Doe"))
            throw new InvalidOperationException("Merged values are missing in the output document.");
    }
}
