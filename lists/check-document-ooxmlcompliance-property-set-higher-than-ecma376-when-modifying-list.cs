using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;

namespace AsposeWordsExamples
{
    public class ListComplianceExample
    {
        public static void Run()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Ensure the document has at least one list.
            List list = doc.Lists.Add(ListTemplate.NumberDefault);
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ListFormat.List = list;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // Modify the first list definition – enable restart numbering at each section.
            list.IsRestartAtEachSection = true;

            // Set OOXML compliance higher than ECMA-376 to persist the property.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Transitional,
                SaveFormat = SaveFormat.Docx
            };

            // Save to a temporary file.
            string outputPath = Path.Combine(Path.GetTempPath(), $"ListCompliance_{Guid.NewGuid()}.docx");
            doc.Save(outputPath, saveOptions);

            // Load the saved document and display its compliance level.
            Document loaded = new Document(outputPath);
            Console.WriteLine($"Document compliance after save: {loaded.Compliance}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ListComplianceExample.Run();
        }
    }
}
