using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace TemplNET
{
    /// <summary>
    /// Represents a reference-by-reflection to a Entry (a field/property/member) within an object. 
    /// The entry's nested Path is defined with a string.
    /// <para/> This is a utility class for Module handlers.
    /// </summary>
    /// <example>
    /// <code>
    /// public class MyContainer { public int x = 5; }
    ///
    /// public class MyModel {
    ///
    ///     [DisplayFormat(ApplyFormatInEditMode = true, DataFormatString = "{0:dd/MM/yy}")]
    ///     public DateTime Date => DateTime.Today.Date
    ///
    ///     public string[] strings = new string[] { "a", "b", "c" };
    ///
    ///     public MyContainer container = new MyContainer();
    /// }
    ///
    /// var model = new MyModel();
    /// 
    /// var entry1 = TemplModelEntry.Get(model, "Date");        // today's date
    /// var entry2 = TemplModelEntry.Get(model, "container.x"); // 5
    /// var entry3 = TemplModelEntry.Get(model, "strings[1]");  //"b"
    /// var entry3keys = entry2.ToStringKeys();                 //"0", "1", "2"
    ///
    /// object o = entry1.Value;
    /// Type t = entry1.Type; // DateTime
    /// DateTime date = <![CDATA[ entry1.AsType<DateTime>(); ]]>
    /// string s = entry1.ToString(); // auto-formats to 'dd/MM/yy' using the DisplayFormat Attribute of MyModel.Date
    ///
    /// </code>
    /// </example>
    public class TemplModelEntry
    {
        /// <summary>
        /// The object containing the Entry
        /// </summary>
        public object Model;
        /// <summary>
        /// A string representing the path to an Entry within the Model. Example: "containerObjectInModel.value"
        /// </summary>
        public string Path;
        /// <summary>
        /// The name of the Entry in the Model. Derived from the Path
        /// </summary>
        public string Name => Path.Split('.').Last();
        /// <summary>
        /// The Field/Property info for the Entry
        /// </summary>
        public MemberInfo Info;
        /// <summary>
        /// The un-typed value of the member at the Entry. Evaluated against the stored Model instance
        /// </summary>
        public object Value;
        /// <summary>
        /// Defines whether the Entry is valid. This requires non-null member info and value for the Entry.
        /// </summary>
        public bool Exists => (Info != null && Value != null);
        /// <summary>
        /// The Type of the member at the Entry
        /// </summary>
        public Type Type => Value.GetType();
        /// <summary>
        /// Extracts typed Attributes from the member info. A common example is DisplayFormatAttribute
        /// </summary>
        /// <typeparam name="T"></typeparam>
        public T[] Attributes<T>() where T : Attribute =>
            Array.ConvertAll(Info?.GetCustomAttributes(typeof(T), true), a => (T)a);
        /// <summary>
        /// Extracts one typed Attribute from the member info. A common example is DisplayFormatAttribute
        /// </summary>
        /// <typeparam name="T"></typeparam>
        public T Attribute<T>() where T : Attribute => Attributes<T>()?.FirstOrDefault();
        /// <summary>
        /// The un-typed value of the member at the Entry. Evaluated against the stored Model instance.
        /// <para/> Throws exception if cannot be cast.
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        /// <typeparam name="T"></typeparam>
        public T AsType<T>()
        {
            if (Value is T)
            {
                return (T)Value;
            }
            throw new InvalidCastException($"Templ: Failed casting model entry \"{Path} \" to \" {typeof(T).Name}\"");
        }

        /// <summary>
        /// Converts the Value for this Entry to string, applying DisplayFormatAttribute if it is defined.
        /// </summary>
        public override string ToString()
        {
            try
            {
                var formatter = Attribute<DisplayFormatAttribute>()?.DataFormatString;
                return (formatter == null ? Value?.ToString() : String.Format(formatter, Value));
            }
            catch (Exception)
            {
                throw new FormatException($"Templ: Failed to convert model property {Path} of type \"{Type}\" to a string");
            }
        }

        /// <summary>
        /// Tries to convert the Value for this Entry to an array[], dict{} or collection, and extract the keys as strings.
        /// <para/> Throws exception if cannot be cast.
        /// </summary>
        /// <exception cref="InvalidCastException"></exception>
        /// <returns></returns>
        public IList<string> ToStringKeys()
        {
            if (Value is object[])
            {
                return Enumerable.Range(0, (Value as object[]).Length).Select(i => i.ToString()).ToList();
            }
            if (Value is IEnumerable<object>)
            {
                return Enumerable.Range(0, (Value as IEnumerable<object>).Count()).Select(i => i.ToString()).ToList();
            }
            if (Value is Dictionary<object, object>)
            {
                return (Value as Dictionary<object, object>).Keys.Select(k => k.ToString()).ToList();
            }
            throw new InvalidCastException($"Templ: Failed to retrieve keys from a collection, dict or array at model path \"{Path}\"; its actual type is \"{Type}\"");
        }

        private TemplModelEntry() { } //Constructor is private

        /// <summary>
        /// Gets a model entry at the specified Path in the given Model.
        /// <para/> Throws exception if cannot be found.
        /// </summary>
        /// <param name="model"></param>
        /// <param name="path"></param>
        public static TemplModelEntry Get(object model, string path)
        {
            var propVal = MemberValue.FindPath(model, path.Trim());
            return new TemplModelEntry()
            {
                Model = model,
                Path = path.Trim(),
                Info = propVal.Info,
                Value = propVal.Value
            };
        }

        /// <summary>
        /// Represents a reference-by-reflection to a member within an object.
        /// <path/> This is a utility class for ModelEntry.
        /// </summary>
        private class MemberValue
        {
            /// <summary>
            /// Matches a collection "index[containing.a.path]"
            /// </summary>
            public static readonly Regex SubPathRegex = new Regex(@"\[[^\[]*\.[^\[]*\]");
            /// <summary>
            /// Matches nested collection "indexes[like[this]]"
            /// </summary>
            public static readonly Regex SubIndexRegex = new Regex(@"\[[^\[]*\[[^\[]*\][^\[]*\]");
            /// <summary>
            /// Matches invalid "[key]path" interaction
            /// </summary>
            public static readonly Regex BadIndex1Regex = new Regex(@"\][^\[\]\.]+\.");
            /// <summary>
            /// Matches invalid "path.[key]" interaction
            /// </summary>
            public static readonly Regex BadIndex2Regex = new Regex(@"\.\[");
            /// <summary>
            /// Matches a collection "index[string]"
            /// </summary>
            public static readonly Regex IndexRegex = new Regex(@"([^\[]+)\[([^\]]+)\]");

            public MemberInfo Info;
            public object Value;

            private MemberValue() { } //Constructor is private
            private MemberValue(MemberInfo info, object value)
            {
                Info = info;
                Value = value;
            }

            /// <summary>
            /// Finds a nested member/value given a model object and string path.
            /// <para/> Throws exception if problems are found.
            /// </summary>
            /// <param name="model">The model.</param>
            /// <param name="path">The path.</param>
            public static MemberValue FindPath(object model, string path)
            {
                CheckPath(path);
                return path.Split('.').Aggregate(new MemberValue() { Value = model }, FindSingle);
            }

            /// <summary>
            /// Performs basic checks on the validity of a string Path.
            /// <para/> Throws exception if problems are found.
            /// </summary>
            /// <param name="path"></param>
            private static void CheckPath(string path)
            {
                // Block use of sub-path as collection index
                if (SubPathRegex.Match(path).Success)
                {
                    throw new NotSupportedException($"Templ: Tried to match a model path \"{path}\" where the collection key is itself a model path. Only int/string keys supported");
                }
                // Block use of nested collection referencing
                if (SubIndexRegex.Match(path).Success)
                {
                    throw new NotSupportedException($"Templ: Tried to match a model path \"{path}\" where the collection key is itself a collection element. Only int/string keys supported");
                }
                // Check for bad use of collection indexing
                if (BadIndex1Regex.Match(path).Success || BadIndex2Regex.Match(path).Success)
                {
                    throw new FormatException($"Templ: Tried to match a model path \"{path}\" with invalid use of collection indexing [].");
                }
            }

            /// <summary>
            /// Uses (dynamic) casting to return an item from a <![CDATA[ Dictionary<string, object> ]]>.
            /// <para/> Throws exception if cannot be cast.
            /// </summary>
            /// <typeparam name="TKey">The type of the key.</typeparam>
            /// <typeparam name="TValue">The type of the value.</typeparam>
            /// <param name="dictionary"></param>
            /// <param name="key"></param>
            /// <example>
            /// <code>
            /// object dict; // some object, may or may not be a dictionary
            /// 
            /// object value = DynamicDict((dynamic)dict, "myStringKey");
            /// </code>
            /// </example>
            private static TValue DynamicDict<TKey, TValue>(Dictionary<TKey, TValue> dictionary, TKey key)
            {
                return dictionary[key];
            }

            /// <summary>
            /// Get a member/value, given an object and member name.
            /// <para/> This is a utility method for FindSingle
            /// </summary>
            /// <param name="obj"></param>
            /// <param name="name"></param>
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
                    throw new FieldAccessException($"Templ: Failed to retrieve property \"{name}\" from model");
                }
            }

            /// <summary>
            /// Get a member/value, given a parent member/value and member name.
            /// </summary>
            /// <param name="parentMember"></param>
            /// <param name="name"></param>
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
                        throw new FieldAccessException($"Templ: Failed to index into model collection/dict/arr \"{name}\"");
                    }
                    return new MemberValue(member.Info, value);
                }
                return Get(parentMember.Value, name);
            }
        }
    }
}