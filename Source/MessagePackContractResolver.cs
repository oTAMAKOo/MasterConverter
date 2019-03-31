﻿
using System;
using MessagePack;
using MessagePack.Formatters;
using MessagePack.Resolvers;
using Modules.MessagePack;

namespace MasterConverter
{
    public class MessagePackContractResolver : IFormatterResolver
    {
        public static IFormatterResolver Instance = new MessagePackContractResolver();

        MessagePackContractResolver() { }

        public IMessagePackFormatter<T> GetFormatter<T>()
        {
            return FormatterCache<T>.formatter;
        }

        private static class FormatterCache<T>
        {
            public static readonly IMessagePackFormatter<T> formatter;

            static FormatterCache()
            {
                IFormatterResolver[] resolvers = null;
                
                resolvers = Resolvers();

                foreach (var item in resolvers)
                {
                    var f = item.GetFormatter<T>();
                    if (f != null)
                    {
                        formatter = f;
                        return;
                    }
                }
            }
        }

        private static IFormatterResolver[] Resolvers()
        {
            return new IFormatterResolver[]
            {
                // DateTime.
                DateTimeResolver.Instance,

                // Builtin.
                BuiltinResolver.Instance,

                // Enum.
                DynamicEnumResolver.Instance,

                // Array, Tuple, Collection.
                DynamicGenericResolver.Instance,

                // Union(Interface).
                DynamicUnionResolver.Instance,
                
                // Object (Map Mode).
                DynamicContractlessObjectResolver.Instance,

                // Primitive.
                PrimitiveObjectResolver.Instance,
            };
        }
    }    
}
