<img alt="Templ.NET" src="https://dl.dropboxusercontent.com/u/39512614/github/templ-dot-net/templ.net.png">

[NuGet Package](https://www.nuget.org/packages/Templ.NET) | [Templates](https://github.com/CPonty/templ-dot-net/tree/master/Examples/ConsoleApp/Templates) | [Example Applications](https://github.com/CPonty/templ-dot-net/tree/master/Examples/ConsoleApp) | [Documentation](#)

<img alt="NugetVersion" src="https://img.shields.io/nuget/v/Templ.NET.svg" />

***

A C# report generation engine, combining .docx templates with strongly-typed data models.

**Templ.NET** is built upon the [DocX library](https://docx.codeplex.com/).

***

#### Install via Nuget
```
Package-Install Templ.NET
```

#### Simple Usage

<img alt="HelloWorld Before" src="https://dl.dropboxusercontent.com/u/39512614/github/templ-dot-net/examples-before.PNG" width="320">
<img alt="HelloWorld After" src="https://dl.dropboxusercontent.com/u/39512614/github/templ-dot-net/examples-after.PNG" width="325">

```C#
var document = Templ.Load("C:\template.docx").Build( new { Title = "Hello World!" });

// Console application
document.SaveAs("C:\output.docx");

// ASP.NET MVC controller 
return new FileContentResult(document.Bytes, Templ.DocxMIMEType);
```

