using System;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace MailMergeExample
{
    // Simple data class for mail merge
    public class Person
    {
        public string Name { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document
            Document doc = new Document();

            // Use DocumentBuilder to add a merge field to the document
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Dear <<Name>>,");
            builder.Writeln("This is a test of Aspose.Words mail merge functionality.");
            builder.Writeln("Best regards,");
            builder.Writeln("Your Company");

            // Prepare data source
            Person person = new Person { Name = "John Doe" };

            // Execute mail merge with the data object
            doc.MailMerge.Execute(new string[] { "Name" }, new object[] { person.Name });

            // Save the merged document as DOCX
            doc.Save("Result.docx");
        }
    }
}
