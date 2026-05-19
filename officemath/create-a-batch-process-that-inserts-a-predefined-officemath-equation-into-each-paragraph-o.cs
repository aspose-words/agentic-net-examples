using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few sample paragraphs.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Define the EQ field argument that will be used for every equation.
        // This creates a simple fraction 1/2.
        const string eqArgument = @"\f(1,2)";

        // Process each paragraph and insert the equation.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph paragraph in paragraphs)
        {
            // Move the builder to the start of the current paragraph.
            builder.MoveTo(paragraph);

            // Insert an EQ field (the field is updated immediately).
            FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

            // Write the EQ argument into the field separator.
            builder.MoveTo(eqField.Separator);
            builder.Write(eqArgument);

            // Return the builder to the paragraph (the field's start parent node).
            builder.MoveTo(eqField.Start.ParentNode);

            // Convert the EQ field to a real OfficeMath object.
            OfficeMath officeMath = eqField.AsOfficeMath();

            if (officeMath != null)
            {
                // Insert the OfficeMath node before the field start and remove the original field.
                eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
                eqField.Remove();

                // Optional: set display type and justification for the top‑level equation.
                officeMath.DisplayType = OfficeMathDisplayType.Display;
                officeMath.Justification = OfficeMathJustification.Left;
            }
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "BatchEquations.docx");
        doc.Save(outputPath);

        // Simple validation – ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
