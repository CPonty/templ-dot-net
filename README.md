![](https://dl.dropboxusercontent.com/u/39512614/github/templ-dot-net/templ-base64.svg)

***

A C# report generation engine, combining .docx templates with strongly-typed data models.

**Templ.NET** is built upon the [DocX library](https://docx.codeplex.com/).

***

#### [Install via Nuget](https://www.nuget.org/packages/Templ.NET)
```
Package-Install Templ.NET
```

#### Simple Usage

```C#
var data = new { title = "Hello World!" } ;
var document = new Templ("C:\template.docx").Build(data);

// Console application
document.SaveAs("C:\output.docx");

// ASP.NET MVC controller 
return new FileContentResult(document.Bytes, Templ.DocxMIMEType);
```

#### Documentation

- Coming soon!
