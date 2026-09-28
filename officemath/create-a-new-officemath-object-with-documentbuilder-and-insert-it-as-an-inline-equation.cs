using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathInlineExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the inline equation.
        builder.Writeln("Inline equation example:");

        // Insert an empty paragraph where the equation will be placed.
        builder.Writeln();

        // Insert an EQ field (field type Equation) with a placeholder for the equation argument.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ eqField = (FieldEQ)field;

        // Write a simple EQ argument (fraction) into the field separator.
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)"); // renders as 1⁄2

        // Update the field so that Aspose.Words can parse the EQ argument.
        eqField.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start node.
            Node fieldStart = eqField.Start;
            Node parent = fieldStart.ParentNode;

            // The parent of a field start is a Paragraph, which derives from CompositeNode.
            if (parent is CompositeNode compositeParent)
            {
                compositeParent.InsertBefore(officeMath, fieldStart);
            }
            else
            {
                throw new InvalidOperationException("Unexpected parent node type for the field start.");
            }

            // Remove the original EQ field, leaving only the OfficeMath node.
            eqField.Remove();
        }
        else
        {
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");
        }

        // Save the document.
        string outputPath = "OfficeMathInline.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
