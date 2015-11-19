using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Novacode;

namespace TemplNET
{
    public abstract class TemplMatch
    {
        public bool Success = false;
        public bool Expired = false;
        public bool Removed = false;
        protected bool RemovedPlaceholder = false;
        public string Body;
        public string[] Fields => Body.Split(':');
        public TemplRegex Regex;
        public string Placeholder => Regex.Text(Body);
        public Regex Pattern => Regex.Pattern;

        // if someone tries to implement single-result search: .DefaultIfEmpty(new T()).First();
        public static IEnumerable<T> Find<T>(TemplRegex rgx, IEnumerable<Paragraph> paragraphs) where T : TemplMatchPara, new()
        {
            return paragraphs.SelectMany(p => Find<T>(rgx, p));
        }
        public static IEnumerable<T> Find<T>(TemplRegex rgx, IEnumerable<Paragraph> paragraphs, int n) where T : TemplMatchPara, new()
        {
            return paragraphs.SelectMany(p => Find<T>(rgx, p).Take(n));
        }
        public static IEnumerable<T> Find<T>(TemplRegex rgx, Paragraph p) where T : TemplMatchPara, new()
        {
            var P = p;
            var t = P.Text;
            return Find<T>(rgx, p.Text).Select(m =>
            {
                m.Paragraph = p;
                return m;
            });
        }
        public static IEnumerable<T> Find<T>(TemplRegex rgx, string s) where T : TemplMatch, new()
        {
            return rgx.Pattern.Matches(s).Cast<Match>().Where(m => m.Success).Select(m => new T()
            {
                Success = true,
                Regex = rgx,
                Body = m.Groups[1].Value,
            });
        }
    }
    public abstract class TemplMatchPara : TemplMatch
    {
        public Paragraph Paragraph;

        private void InsertText(int idx, string text)
        {
            // Simple. Text at index
            Paragraph.InsertText(idx, text);
        }

        public T ToText<T>(string text) where T : TemplMatchPara, new()
        {
            // Insert Text at each match, then remove placeholder
            Paragraph.ReplaceText(Placeholder, text);
            RemovedPlaceholder = true;
            return (T)this;
        }
        public Paragraph AddBefore<T>(string text) where T : TemplMatchPara, new() => Paragraph.InsertParagraphBeforeSelf(text);
        public Paragraph AddAfter<T>(string text) where T : TemplMatchPara, new() => Paragraph.InsertParagraphAfterSelf(text);

        private void InsertPicture(TemplGraphic graphic, int width)
        {
            // Simple. Picture at index(es). Ensure you load picture first.
            Paragraph.FindAll(Placeholder).ForEach(idx =>
            {
                if (width>0)
                {
                    Paragraph.InsertPicture(graphic.Picture((uint)width), idx + Placeholder.Length);
                }
                else
                {
                    Paragraph.InsertPicture(graphic.Picture(), idx + Placeholder.Length);
                }
                Paragraph.Alignment = (graphic.Alignment ?? Paragraph.Alignment);
            });
        }
        public T ToPicture<T>(TemplGraphic graphic, int width) where T : TemplMatchPara, new()
        {
            // Use InsertPicture at each match, then remove placeholder
            if (Removed || RemovedPlaceholder || Expired)
            {
                return (T)this;
            }
            InsertPicture(graphic,width);
            RemovePlaceholder();
            return (T)this;
        }

        public T ToPictures<T>(ICollection<TemplGraphic> graphics, int width) where T : TemplMatchPara, new()
        {
            return ToPictures<T>(graphics.ToArray(), width);
        }

        public T ToPictures<T>(TemplGraphic[] graphics, int width) where T : TemplMatchPara, new()
        {
            if (Removed || RemovedPlaceholder || Expired)
            {
                return (T)this;
            }
            // Use InsertPicture for each graphic, then remove placeholder.
            // Reverse order as pictures get inserted immediately after the placeholder, 
            //  so each subsequent picture inserted will appear before those already inserted.
            foreach (var graphic in graphics.Reverse())
            {
                InsertPicture(graphic, width);
            }
            RemovePlaceholder();
            return (T)this;
        }

        public virtual void Remove()
        {
            // By default, Remove() simply removes the placeholder. 
            // For more advanced content it will do more.
            if (Removed)
            {
                return;
            }
            RemovePlaceholder();
            Removed = true;
        }

        public virtual void RemovePlaceholder()
        {
            if (RemovedPlaceholder)
            {
                return;
            }
            Paragraph.ReplaceText(Placeholder, string.Empty);
            RemovedPlaceholder = true;
        }

        public bool RemoveExpired()
        {
            if (Expired)
            {
                Remove();
            }
            return Removed;
        }
    }
}