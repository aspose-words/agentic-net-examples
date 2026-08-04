using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

public class Program
{
    // Simple model that matches the JSON structure.
    private class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public static void Main()
    {
        // 1. Prepare a JSON array with sample data.
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Charlie"", ""Age"": 28 }
        ]";

        // 2. Deserialize the JSON into a list of Person objects.
        List<Person> people = JsonConvert.DeserializeObject<List<Person>>(json) ?? new List<Person>();

        // 3. Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Optional: add a title before the repeating section.
        builder.Writeln("People List:");
        builder.Writeln(); // empty line for spacing.

        // 4. Create a block‑level repeating section content control.
        StructuredDocumentTag repeatingSection = new StructuredDocumentTag(
            doc,
            SdtType.RepeatingSection,
            MarkupLevel.Block)
        {
            Title = "PeopleRepeatingSection",
            Tag = "people-section"
        };

        // 5. For each person, create a paragraph and add it to the repeating section.
        foreach (Person person in people)
        {
            Paragraph para = new Paragraph(doc);
            // The paragraph text can be formatted as needed.
            para.AppendChild(new Run(doc, $"Name: {person.Name}, Age: {person.Age}"));
            repeatingSection.AppendChild(para);
        }

        // 6. Insert the repeating section into the document body.
        doc.FirstSection.Body.AppendChild(repeatingSection);

        // 7. Save the resulting document.
        const string outputPath = "RepeatingSectionFromJson.docx";
        doc.Save(outputPath);
    }
}
