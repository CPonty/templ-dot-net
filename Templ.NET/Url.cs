namespace TemplNET
{
    /// <summary>
    /// Used to load/store named URLs and generate Hyperlinks for the document
    /// </summary>
    /// <example>
    /// <code>
    /// using Novacode;
    ///
    /// var url = new TemplUrl(){Url = "http://www.google.com", Text = "Google"};
    /// 
    /// // Include link in the model
    /// Templ.Load("C:\template.docx")
    ///      .Build( new { Title = "Hello World!", weblink = url})
    ///      .SaveAs("C:\output.docx");
    /// </code>
    /// </example>
    public class TemplUrl
    {
        public string Text;
        public string Url;
        public bool ToDelete = false;
    }
}