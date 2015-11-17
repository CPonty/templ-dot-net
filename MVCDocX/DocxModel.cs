using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Web;

namespace DocXMVC
{
    public class DocxModelEntry
    {
        public object Model;
        public string Path;
        public string Name => Path.Split('.').Last();
        public MemberInfo Info;
        public object Value;
        public bool Exists => (Info != null && Value != null);
        public Type Type => Value.GetType();
        public T AsType<T>()
        {
            if (Value is T)
            {
                return (T)Value;
            }
            throw new InvalidCastException($"DocX: Failed casting model entry \"{Path} \" to \" {typeof(T).Name}\"");
        }
        public T[] Attributes<T>() where T : Attribute
        {
            return Array.ConvertAll(Info?.GetCustomAttributes(typeof(T), true), a => (T)a);
        }
        public T Attribute<T>() where T : Attribute
        {
            return Attributes<T>()?.FirstOrDefault();
        }
        public override string ToString()
        {
            try
            {
                var formatter = Attribute<DisplayFormatAttribute>()?.DataFormatString;
                return (formatter == null ? Value?.ToString() : String.Format(formatter, Value));
            }
            catch (Exception)
            {
                throw new FormatException($"DocX: Failed to convert model property {Path} of type \"{Type}\" to a string");
            }
        }
        public IList<string> ToStringKeys()
        {
            if (Value is object[])
            {
                return Enumerable.Range(0, (Value as object[]).Length).Select(i => i.ToString()).ToList();
            }
            if (Value is ICollection<object>)
            {
                return Enumerable.Range(0, (Value as ICollection<object>).Count).Select(i => i.ToString()).ToList();
            }
            if (Value is Dictionary<object, object>)
            {
                return (Value as Dictionary<object, object>).Keys.Select(k => k.ToString()).ToList();
            }
            throw new InvalidCastException($"Docx: Failed to retrieve keys from a collection, dict or array at model path \"{Path}\"; its actual type is \"{Type}\"");
        }

        private DocxModelEntry() { }
        public static DocxModelEntry Get(object model, string path)
        {
            var propVal = MemberValue.FindPath(model, path);
            return new DocxModelEntry()
            {
                Model = model,
                Path = path,
                Info = propVal.Info,
                Value = propVal.Value
            };
        }
        private class MemberValue
        {
            public static readonly Regex SubPathRegex = new Regex(@"\[[^\[]*\.[^\[]*\]");
            public static readonly Regex SubIndexRegex = new Regex(@"\[[^\[]*\[[^\[]*\][^\[]*\]");
            public static readonly Regex BadIndex1Regex = new Regex(@"\][^\[\]\.]+\.");
            public static readonly Regex BadIndex2Regex = new Regex(@"\.\[");
            public static readonly Regex IndexRegex = new Regex(@"([^\[]+)\[([^\]]+)\]");


            public MemberInfo Info;
            public object Value;

            private MemberValue() { }
            private MemberValue(MemberInfo info, object value)
            {
                Info = info;
                Value = value;
            }
            public static MemberValue FindPath(object model, string path)
            {
                CheckPath(path);
                return path.Split('.').Aggregate(new MemberValue() { Value = model }, FindSingle);
            }

            private static void CheckPath(string path)
            {
                // Block use of sub-path as collection index
                if (SubPathRegex.Match(path).Success)
                {
                    throw new NotSupportedException("DocX: Tried to match a model path \"" + path + "\" where the collection key is itself a model path. Only int/string keys supported");
                }
                // Block use of nested collection referencing
                if (SubIndexRegex.Match(path).Success)
                {
                    throw new NotSupportedException("DocX: Tried to match a model path \"" + path + "\" where the collection key is itself a collection element. Only int/string keys supported");
                }
                // Check for bad use of collection indexing
                if (BadIndex1Regex.Match(path).Success || BadIndex2Regex.Match(path).Success)
                {
                    throw new FormatException("DocX: Tried to match a model path \"" + path + "\" with invalid use of collection indexing [].");
                }
            }
            private static TValue DynamicDict<TKey, TValue>(Dictionary<TKey, TValue> dictionary, TKey key)
            {
                return dictionary[key];
            }
            private static MemberValue Get(object obj, string name)
            {
                try
                {
                    Type type = obj.GetType();
                    object value = null;
                    MemberInfo member = null;
                    try
                    {
                        // Try retrieving a property
                        var property = type.GetProperty(name);
                        value = property.GetValue(obj, null);
                        member = property as MemberInfo;
                    }
                    catch
                    {
                        // If property fails, try retrieving a field.
                        // If field fails, it will be caught by the outer "catch-throw FieldAccessException"
                        var field = type.GetField(name);
                        value = field.GetValue(obj);
                        member = field as MemberInfo;
                    }
                    return new MemberValue(member, value);
                }
                catch
                {
                    throw new FieldAccessException($"DocX: Failed to retrieve property \"{name}\" from model");
                }
            }
            private static MemberValue FindSingle(MemberValue parentMember, string name)
            {
                if (IndexRegex.IsMatch(name))
                {
                    var groups = IndexRegex.Match(name).Groups;
                    var member = Get(parentMember.Value, groups[1].Value);
                    var key = groups[2].Value;
                    object value = null;
                    try
                    {
                        // Assume it's a dict
                        value = DynamicDict((dynamic)member.Value, key);
                    }
                    catch
                    {
                        try
                        {
                            // Assume it's enumerable
                            var enumerable = member.Value as IEnumerable<object>;
                            value = enumerable?.ElementAt(int.Parse(key));
                        }
                        catch
                        {
                            try
                            {
                                // Assume it's an array
                                var array = member.Value as object[];
                                value = array[int.Parse(key)];
                            }
                            catch { }
                        }
                    }
                    // If value is null, either all the try-catches failed or the field truly is null. Either way, no-go.
                    if (value == null)
                    {
                        throw new FieldAccessException($"DocX: Failed to index into model collection/dict/arr \"{name}\"");
                    }
                    return new MemberValue(member.Info, value);
                }
                return Get(parentMember.Value, name);
            }
        }
    }
}