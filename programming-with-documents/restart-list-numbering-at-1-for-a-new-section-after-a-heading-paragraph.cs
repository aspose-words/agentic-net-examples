using System;
using Aspose.Words;
using Aspose.Words.Lists;

namespace RestartListNumberingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert first heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Section 1");

            // Create the first numbered list (will start at 1).
            List firstList = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = firstList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            // End list formatting for the next paragraph.
            builder.ListFormat.RemoveNumbers();

            // Insert second heading.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Section 2");

            // Create a new numbered list – numbering restarts at 1.
            List secondList = doc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = secondList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.RemoveNumbers();

            // Save the document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);
        }
    }
}
