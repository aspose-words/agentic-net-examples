using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a simple template with mail merge fields.
        string template = @"{""Document"":{""Body"":{""Paragraphs"":[{""Runs"":[{""Text"":""Hello <<Name>>, today is <<Date>>.""}]}]}}";

        // Load the template into an Aspose.Words Document.
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(template)))
        {
            Document doc = new Document(stream);

            // Perform mail merge with sample data.
            string[] fieldNames = { "Name", "Date" };
            object[] fieldValues = { "John Doe", DateTime.Now.ToString("D") };
            doc.MailMerge.Execute(fieldNames, fieldValues);

            // Save the merged document as DOCX.
            doc.Save("MergedDocument.docx", SaveFormat.Docx);
        }
    }
}
