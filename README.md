# Netphi
Ever wondered on how to combine the strengths of both Delphi and .NET to create powerful, versatile applications?

**Disclaimer:** In the following, I will only show how to use .NET in Delphi and not the other way around.

 # Godzilla-type breaking change
According to the [official documentation](https://learn.microsoft.com/en-us/dotnet/core/native-interop/expose-components-to-com#embedding-type-libraries-in-the-com-host)...

> Unlike in .NET Framework, there is no support in .NET Core or .NET 5+ for generating a COM Type Library (TLB) from a .NET assembly. The guidance is to either manually write an IDL file or a C/C++ header for the native declarations of the COM interfaces. If you decide to write an IDL file, you can compile it with the Visual C++ SDK's MIDL compiler to produce a TLB.

... there is no support in .NET Core or .NET 5+ for generating a typelib, which is needed to use .NET in Delphi.

## What about .NET Framework?
While many projects have already migrated to .NET Core or .NET 5+, there are still some applications running on the .NET Framework. If you're one of them and looking to use .NET in Delphi, you're in luck. Simply follow [this tutorial](https://www.youtube.com/watch?v=ZutlhThQJ5s&ab_channel=AliY%C4%B1ld%C4%B1r%C4%B1m) on YouTube.

# Expose .NET components to COM
The following steps can be also found in [this official documentation](https://learn.microsoft.com/en-us/dotnet/core/native-interop/expose-components-to-com#embedding-type-libraries-in-the-com-host).

## Create or change an existing library
```csharp
using System;
using System.Runtime.InteropServices;

namespace Netphi;

[ComVisible(true)]
[Guid("B8E75073-0EEE-41C8-86EB-8B30AFE5C0EC")]
public sealed class Example
{
    public string GetHelloWorld()
    {
        return "Hello World!";
    }
}
```

## Change your .csproj
```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net7.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <EnableComHosting>true</EnableComHosting>
  </PropertyGroup>

</Project>
```
After compiling the project you will have a `ProjectName.dll`, `ProjectName.deps.json`, `ProjectName.runtimeconfig.json` and `ProjectName.comhost.dll` file.

## Register the COM host for COM
Open an elevated command prompt and run `regsvr32 ProjectName.comhost.dll`. That will register all of your exposed .NET objects with COM.

## Reg-Free COM
``
dotnet build src/Netphi.dscom/Netphi.dscom.csproj /p:RegFree=True
``

# Generate the .tlb
[dscom](https://github.com/dspace-group/dscom)
```console
C:\> dotnet tool install dscom -g

C:\> dscom tlbexport myassembly.dll
```

# Useful links
- https://stackoverflow.com/questions/71378349/how-do-i-call-a-net-core-6-0-dll-in-delphi
- https://learn.microsoft.com/en-us/dotnet/core/native-interop/expose-components-to-com
- https://github.com/dotnet/samples/tree/main/core/extensions/COMServerDemo
- https://www.youtube.com/watch?v=ZutlhThQJ5s&ab_channel=AliY%C4%B1ld%C4%B1r%C4%B1m
- https://github.com/dspace-group/dscom
