using System;
using System.Collections.Generic;
using System.Linq;
// ReSharper disable ArgumentsStyleOther
// ReSharper disable ArgumentsStyleNamedExpression

namespace ConsoleTest.Models;

public record Product(
    int Id, 
    string Name, 
    decimal Price, 
    string Description, 
    IEnumerable<Product.Feature> Features)
{
    public record Feature(int Id, string Name, string Description);
    public static IEnumerable<Product> Test(int Count = 10) => Enumerable
       .Range(1, Count)
       .Select(i => new Product(
            Id: i, 
            Name: $"Product-{i}", 
            Price: i * 1000 - i * 10, 
            Description: $"Description of product {i}",
            Features: Enumerable
               .Range(1, Random.Shared.Next(10))
               .Select(n => new Feature(
                    Id: i * Count + n,
                    Name: $"Feature-{i}-{n}",
                    Description: $"Feature.Description-{i}-{n}"))));
}

