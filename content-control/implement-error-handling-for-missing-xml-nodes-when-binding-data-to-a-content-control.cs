using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlXmlBindingExample
{
    // Custom exception to indicate a missing XML node for a content control.
    public class MissingXmlNodeException : Exception
    {
        public MissingXmlNodeException(string message) : base(message) { }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // -----------------------------------------------------------------
            // 1. Create a custom XML part that will serve as the data source.
            // -----------------------------------------------------------------
            string xmlPartId = Guid.NewGuid().ToString("B");
            string xmlContent =
                "<root>" +
                "  <person>" +
                "    <name>John Doe</name>" +
                "    <!-- Note: <age> element is intentionally omitted to simulate a missing node -->" +
                "  </person>" +
                "</root>";

            CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

            // -----------------------------------------------------------------
            // 2. Insert a content control bound to an existing XML node (name).
            // -----------------------------------------------------------------
            builder.Writeln("Name:");
            StructuredDocumentTag nameSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
            nameSdt.Title = "NameControl";
            nameSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]/name[1]", string.Empty);

            // -----------------------------------------------------------------
            // 3. Insert a content control bound to a missing XML node (age).
            // -----------------------------------------------------------------
            builder.Writeln("Age:");
            StructuredDocumentTag ageSdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
            ageSdt.Title = "AgeControl";
            ageSdt.XmlMapping.SetMapping(xmlPart, "/root[1]/person[1]/age[1]", string.Empty);

            // -----------------------------------------------------------------
            // 4. Validate the bindings and handle missing XML nodes.
            // -----------------------------------------------------------------
            foreach (Node node in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
            {
                StructuredDocumentTag sdt = (StructuredDocumentTag)node;

                // Process only those SDTs that have an XML mapping.
                if (sdt.XmlMapping.IsMapped)
                {
                    try
                    {
                        // Retrieve the displayed text of the content control.
                        string displayedText = sdt.GetText().Trim();

                        // If the text is empty, the mapped XML node does not exist.
                        if (string.IsNullOrEmpty(displayedText))
                        {
                            throw new MissingXmlNodeException(
                                $"The XML node referenced by content control '{sdt.Title}' was not found.");
                        }

                        // For demonstration, write the successful binding to the console.
                        Console.WriteLine($"Content control '{sdt.Title}' bound successfully. Value: '{displayedText}'.");
                    }
                    catch (MissingXmlNodeException ex)
                    {
                        // Handle the missing node scenario gracefully.
                        Console.WriteLine($"Error: {ex.Message}");
                    }
                }
            }

            // -----------------------------------------------------------------
            // 5. Save the resulting document.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);
        }
    }
}
