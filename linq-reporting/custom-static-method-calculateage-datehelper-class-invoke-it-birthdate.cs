using System;

public static class DateHelper
{
    // Calculates the age in full years based on the provided birth date.
    public static int CalculateAge(DateTime birthDate)
    {
        // Use today's date for the calculation.
        DateTime today = DateTime.Today;

        // Compute the preliminary age.
        int age = today.Year - birthDate.Year;

        // If the birthday hasn't occurred yet this year, subtract one.
        if (birthDate > today.AddYears(-age))
            age--;

        return age;
    }
}

public class Program
{
    public static void Main()
    {
        // Example usage:
        // Replace the literal below with the actual birth date you need.
        DateTime birthDate = new DateTime(1990, 5, 23);
        int age = DateHelper.CalculateAge(birthDate);
        Console.WriteLine($"Age: {age}");
    }
}
