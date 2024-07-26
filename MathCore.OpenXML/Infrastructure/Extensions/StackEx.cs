namespace MathCore.OpenXML.Infrastructure.Extensions;

internal static class StackEx
{
    public static IEnumerable<T> EnumerateWhileNotEmpty<T>(this Stack<T> stack)
    {
        while(stack.Count > 0)
            yield return stack.Pop();
    }
}
