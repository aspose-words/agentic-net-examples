using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Lists;

namespace AsposeWordsExamples
{
    public class ListComplianceExample
    {
        public static void Run()
        {
            // Load an existing DOCX document.
            // (Replace with the actual path to your source document.)
            string inputPath = @"C:\Docs\InputDocument.docx";
            Document doc = new Document(inputPath);

            // Ensure the document has at least one list.
            if (doc.Lists.Count == 0)
            {
                // If no list exists, create a simple numbered list as a placeholder.
                List newList = doc.Lists.Add(ListTemplate.NumberDefault);
                // Add a couple of paragraphs to demonstrate the list.
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.ListFormat.List = newList;
                builder.Writeln("Item 1");
                builder.Writeln("Item 2");
                builder.ListFormat.RemoveNumbers();
            }

            // Modify the first list definition – for example, enable restart numbering at each section.
            List firstList = doc.Lists[0];
            firstList.IsRestartAtEachSection = true;

            // The IsRestartAtEachSection property is only written to DOCX when the OOXML compliance
            // level is higher than Ecma376_2006. Therefore we create OoxmlSaveOptions with a higher compliance.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                // Choose a compliance level newer than ECMA-376.
                // Iso29500_2008_Transitional is sufficient for the restart-at-each-section feature.
                Compliance = OoxmlCompliance.Iso29500_2008_Transitional,
                SaveFormat = SaveFormat.Docx
            };

            // Save the modified document using the specified compliance level.
            // (Replace with the desired output path.)
            string outputPath = @"C:\Docs\OutputDocument.docx";
            doc.Save(outputPath, saveOptions);

            // Optional: Verify that the saved document reports the expected compliance.
            Document loaded = new Document(outputPath);
            Console.WriteLine($"Document compliance after save: {loaded.Compliance}");
            // Expected output: Iso29500_2008_Transitional
        }
    }

    // Entry point required by the C# compiler.
    public class Program
    {
        public static void Main(string[] args)
        {
            ListComplianceExample.Run();
        }
    }
}
