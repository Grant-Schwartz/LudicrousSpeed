using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("LudicrousSpeed Outlook Attachment Guard")]
[assembly: AssemblyDescription("Warns before a workbook with LudicrousSpeed live cells is mailed out.")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("LudicrousSpeed")]
[assembly: AssemblyProduct("LudicrousSpeed")]
[assembly: AssemblyCopyright("")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Off at assembly level, on for the one class Outlook has to create. Exposing
// every type to COM would put internals in the registry for no reason.
[assembly: ComVisible(false)]
[assembly: Guid("32f2de62-46cf-4a9a-a968-a529b712437d")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyFileVersion("0.1.0.0")]
