using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class SerializeContentControlMappings
{
    static void Main()
    {
        // Load the source document that contains content controls (structured document tags).
        // The Document constructor is the provided load rule.
        Document doc = new Document("InputDocument.docx");

        // Prepare a StringBuilder to collect all XML schema definitions
        // associated with the custom XML parts referenced by the content controls.
        StringBuilder schemaBuilder = new StringBuilder();

        // Iterate through every custom XML part in the document.
        foreach (CustomXmlPart xmlPart in doc.CustomXmlParts)
        {
            // Each CustomXmlPart has a Schemas collection (type CustomXmlSchemaCollection).
            // Add every schema string to the output.
            foreach (string schema in xmlPart.Schemas)
            {
                schemaBuilder.AppendLine(schema);
            }
        }

        // Define the path for the external XSD file that will store the collected schemas.
        string xsdPath = "ContentControlMappings.xsd";

        // Write the accumulated schema definitions to the XSD file.
        // The File.WriteAllText method follows the provided save rule (writing to external storage).
        File.WriteAllText(xsdPath, schemaBuilder.ToString());

        // Optional: confirm that the file has been created.
        Console.WriteLine($"XML schema mapping saved to: {Path.GetFullPath(xsdPath)}");
    }
}
