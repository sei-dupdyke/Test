using System;

class Program
{
    /// <summary>
    /// Entry point for the application. Asks the user for two numbers,
    /// uses <see cref="Adder"/> to add them and prints the result.
    /// </summary>
    static void Main(string[] args)
    {
        Console.WriteLine("Enter first number:");
        if (!int.TryParse(Console.ReadLine(), out int first))
        {
            Console.WriteLine("Invalid input");
            return;
        }

        Console.WriteLine("Enter second number:");
        if (!int.TryParse(Console.ReadLine(), out int second))
        {
            Console.WriteLine("Invalid input");
            return;
        }

        var adder = new Adder();
        int result = adder.Add(first, second);
        Console.WriteLine($"The sum is: {result}");
    }
}
