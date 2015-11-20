<img alt="Templ.NET" src="https://dl.dropboxusercontent.com/u/39512614/github/templ-dot-net/templ.net.png" width="640">

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
public class HelloWorldModel
{
  public string Title = "Hello World!";
}

var data = new HelloWorldModel();
var document = Templ.Load("C:\template.docx").Build(data);

// Console application
document.SaveAs("C:\output.docx");

// ASP.NET MVC controller 
return new FileContentResult(document.Bytes, Templ.DocxMIMEType);
```

<img alt="HelloWorld Before" src="https://dl.dropboxusercontent.com/u/39512614/github/templ-dot-net/examples-before.PNG" width="320">
<img alt="HelloWorld After" src="https://dl.dropboxusercontent.com/u/39512614/github/templ-dot-net/examples-after.PNG" width="325">

#### Documentation

- Coming soon!

#### [Examples](https://github.com/CPonty/templ-dot-net/tree/master/Examples/ConsoleApp)
