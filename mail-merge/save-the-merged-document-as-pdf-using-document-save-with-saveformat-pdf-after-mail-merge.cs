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

        // Build the template with merge fields.
        builder.Writeln("Hello, ");
        builder.InsertField("MERGEFIELD FirstName");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName");
        builder.Writeln("!");

        // Prepare simple data source for the merge.
        string[] fieldNames = new string[] { "FirstName", "LastName" };
        object[] fieldValues = new object[] { "John", "Doe" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Validate that the merge replaced the fields.
        string mergedText = doc.GetText();
        if (!mergedText.Contains("John") || !mergedText.Contains("Doe"))
        {
            throw new InvalidOperationException("Mail merge failed: merged values not found in the document.");
        }

        // Save the merged document as PDF.
        string outputPath = "MergedOutput.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Ensure the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("Failed to create the PDF output file.", outputPath);
        }
    }
}
