using System.Collections.Generic;

namespace System.Linq;

internal static class GroupingExtensions
{
    public static void Deconstruct<TKey, TValue>(this IGrouping<TKey, TValue> Group, out TKey Key, out IEnumerable<TValue> Value)
    {
        Key = Group.Key;
        Value = Group;
    }
}
