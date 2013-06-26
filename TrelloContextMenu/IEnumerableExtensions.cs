using System;
using System.Collections.Generic;

namespace TrelloContextMenu
{
    public static class IEnumerableExtensions
    {
        public static void ForEach<T>(this IEnumerable<T> e, Action<T> action)
        {
            foreach (T item in e)
                action(item);
        }
    }
}