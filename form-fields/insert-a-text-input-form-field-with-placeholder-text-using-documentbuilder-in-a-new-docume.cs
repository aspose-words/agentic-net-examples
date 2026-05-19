using System;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeWordsFormFieldExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a prompt before the form field.
            builder.Write("Please enter your name: ");

            // Insert a text input form field.
            // Parameters: name, type, format, placeholder text, maxLength (0 = unlimited).
            builder.InsertTextInput("NameField", TextFormFieldType.Regular, "", "Enter your name here", 0);

            // Save the document.
            doc.Save("FormFieldTextInput.docx");
        }
    }
}
