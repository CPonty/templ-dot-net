using System;
using System.Collections.Generic;
using System.Linq;
using Novacode;

namespace TemplNET
{
    public abstract class TemplModule<T> : TemplModule where T : TemplMatchPara, new()
    {
        public uint MinFields = 1;
        public uint MaxFields = 1;
        public new Func<T, DocX, object, T> CustomHandler = (m, doc, model) => m;

        public TemplModule(String name, TemplBuilder docBuilder, string[] prefixes)
            : base(name, docBuilder, prefixes)
        { }

        private T CheckFieldCount(T m)
        {
            var l = m.Fields.Length;
            if (l < MinFields)
            {
                throw new FormatException($"Templ: Module \"{GetType()}\" found a placeholder \"{m.Placeholder}\" with too few \":\"-separated fields ({l}, minimum is {MinFields})");
            }
            if (l > MaxFields)
            {
                throw new FormatException($"Templ: Module \"{GetType()}\" found a placeholder \"{m.Placeholder}\" with too many \":\"-separated fields ({l}; maxmimum is {MaxFields})");
            }
            return m;
        }
        public void BuildFromScope(IEnumerable<T> scope)
        {
            // Note how we are constantly "committing" the changes using ToList().
            // This ensures order is preserved (e.g. all "finds" happen before all "handler"s)
            scope.Select( m => CheckFieldCount(m) ).ToList()
                 .Select( m => Handler(m)).ToList()
                 .Where(  m =>!m.RemoveExpired()).ToList()
                 .Select( m => CustomHandler(m, Doc, Model)).ToList()
                 .ForEach(m => m.RemoveExpired());
        }
        public override void Build()
        {
            BuildFromScope(Regexes.SelectMany(rxp => FindAll(rxp)));
        }
        /// <summary>
        /// Container-specific handler.
        /// Modify the underlying container; mark expired to delete
        /// </summary>
        /// <param name="m"></param>
        public abstract T Handler(T m);

        /// <summary>
        /// Container-specific finder.
        /// Get all instances from the document.
        /// </summary>
        /// <param name="rxp"></param>
        public abstract IEnumerable<T> FindAll(TemplRegex rxp);
    }

    public abstract class TemplModule
    {
        public string Name;
        protected TemplBuilder DocBuilder;
        public DocX Doc => DocBuilder.Doc;
        public object Model => DocBuilder.Model;
        public TemplRegex[] Regexes;
        public Func<object, TemplBuilder, object> CustomHandler;

        public TemplModule(string name, TemplBuilder docBuilder, string[] prefixes)
        {
            foreach (var prefix in prefixes.Where(s => s.Contains(":")))
            {
                throw new FormatException($"Templ: Module \"{name}\": prefix \"{prefix}\" cannot contain the split character ':'");
            }
            Regexes = prefixes.Select(pre => new TemplRegex(pre)).ToArray();
            DocBuilder = docBuilder;
            Name = name;
        }

        public abstract void Build();
    }
}