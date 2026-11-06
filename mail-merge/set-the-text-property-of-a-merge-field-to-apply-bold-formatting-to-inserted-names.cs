using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for template creation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple sentence with a merge field for the name.
        builder.Writeln("Dear ");
        builder.InsertField("MERGEFIELD Name");
        builder.Writeln(",");

        // Prepare mail merge data.
        string[] fieldNames = { "Name" };
        object[] fieldValues = { "John Doe" };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // After merging, locate the run that contains the inserted name.
        Run nameRun = doc.GetChildNodes(NodeType.Run, true)
                         .Cast<Run>()
                         .FirstOrDefault(r => r.Text == "John Doe");

        if (nameRun == null)
            throw new InvalidOperationException("Merged name not found in the document.");

        // Apply bold formatting to the inserted name.
        nameRun.Font.Bold = true;

        // Save the merged document.
        string outputPath = "MergedBold.docx";
        doc.Save(outputPath);

        // Validation: ensure the file was created and contains bold text.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not created.", outputPath);

        // Load the saved document to verify formatting.
        Document resultDoc = new Document(outputPath);
        Run resultRun = resultDoc.GetChildNodes(NodeType.Run, true)
                                 .Cast<Run>()
                                 .FirstOrDefault(r => r.Text == "John Doe");

        if (resultRun == null || !resultRun.Font.Bold)
            throw new InvalidOperationException("Bold formatting was not applied to the merged name.");
    }
}
