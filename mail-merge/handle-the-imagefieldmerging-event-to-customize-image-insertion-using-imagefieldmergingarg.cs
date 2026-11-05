using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add a merge field for an image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Below is the merged image:");
        builder.InsertField("MERGEFIELD ImageField \\* MERGEFORMAT");

        // Prepare a simple PNG image (a red dot) as a byte array.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFklEQVQoU2NkYGD4z0AEYBxVSFIAAQAA//8DAF6cB6cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Execute mail merge. Provide the image bytes directly as the field value.
        string[] fieldNames = { "ImageField" };
        object[] fieldValues = { imageBytes };
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the merged document.
        string outputPath = "MergedDocument.docx";
        doc.Save(outputPath);

        // Validate that the output file was created and is not empty.
        FileInfo info = new FileInfo(outputPath);
        if (!info.Exists || info.Length == 0)
        {
            throw new InvalidOperationException("The merged document was not created correctly.");
        }

        Console.WriteLine($"Document merged successfully. Output file: {outputPath}");
    }
}
