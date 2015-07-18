namespace ExcelBenchmark.Tests.Extensions
{
    using System.Collections.Generic;

    public static class CollectionExtensions
    {
        public static IList<dynamic> Duplicate(this IList<dynamic> items, int times)
        {
            var results = new List<dynamic>();
            results.AddRange(items);
            for (int i = 0; i < times; i++)
            {
                results.AddRange(items);
            }

            return results;
        }
    }
}