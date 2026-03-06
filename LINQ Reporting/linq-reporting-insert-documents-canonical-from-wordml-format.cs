using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data class that will be used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // WORDML (WordprocessingML) template as a string.
            // The template contains a reporting tag that will be replaced by the data source.
            string wordmlTemplate = @"
<?xml version=""1.0"" encoding=""UTF-8""?>
<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r><w:t>Report for: </w:t></w:r>
      <w:r><w:t><<[person.Name]>></w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Age: </w:t></w:r>
      <w:r><w:t><<[person.Age]>></w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>";

            // Load the WORDML template into an Aspose.Words Document.
            // Use the Document constructor that accepts a Stream (lifecycle rule).
            using (MemoryStream templateStream = new MemoryStream(Encoding.UTF8.GetBytes(wordmlTemplate)))
            {
                Document doc = new Document(templateStream);

                // Prepare the data source.
                Person person = new Person { Name = "John Doe", Age = 42 };

                // Create the reporting engine and populate the template.
                // BuildReport is the method that merges the data with the template.
                ReportingEngine engine = new ReportingEngine();
                // The data source name "person" matches the tag used in the template.
                engine.BuildReport(doc, person, "person");

                // Save the resulting document to a file.
                // Use the Save(string) method (lifecycle rule).
                doc.Save("ReportFromWordML.docx");
            }
        }
    }
}
