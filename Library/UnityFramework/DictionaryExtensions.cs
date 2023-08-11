
using System;
using System.Collections.Generic;

namespace Extensions
{
    public static partial class DictionaryExtensions
    {
        #if !NET6_0_OR_GREATER
        
        public static TValue GetValueOrDefault<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue = default(TValue))
        {
            TValue result;

            return dictionary.TryGetValue(key, out result) ? result : defaultValue;
        }

        #endif

        public static TValue GetOrAdd<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TKey key, Func<TKey, TValue> valueFactory)
        {
            TValue value;

            if (!dictionary.TryGetValue(key, out value))
            {
                value = valueFactory(key);
                dictionary.Add(key, value);
            }

            return value;
        }
    }
}
