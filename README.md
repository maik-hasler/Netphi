# Netphi
This article walks you through how to expose a class to COM from .NET Core (or .NET 5+).

# Table of contents
- [Introduction](#netphi)
- [Table of contents](#table-of-contents)
- [Godzilla-type breaking change](#godzilla-type-breaking-change)
    - [What about .NET Framework?](#what-about-net-framework)
- [Expose .NET Core components to COM](#expose-net-core-components-to-com)
    - [Create the library](#create-the-library)
    - [Generate the COM host](#generate-the-com-host)
    - [Register the COM host for COM](#register-the-com-host-for-com)
- [Generate the type library](#generate-the-type-library)

# Godzilla-type breaking change
According to the [official Microsoft documentation](https://learn.microsoft.com/en-us/dotnet/core/native-interop/expose-components-to-com#embedding-type-libraries-in-the-com-host)...

> Unlike in .NET Framework, there is no support in .NET Core or .NET 5+ for generating a COM Type Library (TLB) from a .NET assembly. The guidance is to either manually write an IDL file or a C/C++ header for the native declarations of the COM interfaces. If you decide to write an IDL file, you can compile it with the Visual C++ SDK's MIDL compiler to produce a TLB.

... there is no support in .NET Core or .NET 5+ to automatically generate a `.tlb` file, which is later needed for the generation of the COM client in Delphi.

## What about .NET Framework?
While many projects have already migrated to .NET Core or .NET 5+, there are still some applications running on the .NET Framework. 
If you're one of them and looking to use .NET in Delphi, you're in luck. Simply follow [this YouTube tutorial](https://www.youtube.com/watch?v=ZutlhThQJ5s&ab_channel=AliY%C4%B1ld%C4%B1r%C4%B1m).

# Expose .NET Core components to COM
The  [official Microsoft documentation](https://learn.microsoft.com/en-us/dotnet/core/native-interop/expose-components-to-com) already provides a guidance on creating and registering the COM host for .NET components. However, there is are multiple potential pitfalls that aren't covered.

## Create the library
The first step is to create the library.
```cs
[ComVisible(true)]
[Guid("A028C948-7924-45EB-B679-6484818A8585")]
public interface IServer
{
    double ComputePi();
}
```
and the implementation
```cs
[ComVisible(true)]
[Guid("D34D391D-5138-48E8-9D3C-6E10002A253E")]
[ComDefaultInterface(typeof(IServer))]
public class Server : IServer
{
    public double ComputePi()
    {
        return Math.PI;
    }
}
```

## Generate the COM host
To enable COM hosting for your project, follow these steps:
1. Open the `.csproj` project file and insert `<EnableComHosting>true</EnableComHosting>` within a `<PropertyGroup></PropertyGroup>` tag.
2. Depending on the architecture of the COM client, execute one of the following commands to build your project:
- For x86 architecture:
```
dotnet build -r win-x86 --no-self-contained .\path\to\ProjectName.csproj
```
- For x64 architecture:
```
dotnet build -r win-x64 --no-self-contained .\path\to\ProjectName.csproj
```
Upon successful execution, you'll find the following files in the output:
- `ProjectName.dll`
- `ProjectName.deps.json`
- `ProjectName.runtimeconfig.json`
- `ProjectName.comhost.dll`

## Register the COM host for COM
Open an **elevated** command prompt and run `regsvr32 ProjectName.comhost.dll`. That will register all of your exposed .NET objects with COM.

<!---
**TODO: Clarify if regsvr32 automatically checks if the ProjectName.comhost.dll is x86 or x64 and therefore register it in the correct registry. Otherwise use this command:**
For x86:
```
%windir%\SysWoW64\regsvr32.exe %windir%\SysWoW64\ProjectName.comhost.dll
```
For x64:
```
%windir%\System32\regsvr32.exe %windir%\System32\ProjectName.comhost.dll
```
-->

# Generate the type library
While there are many ways to generate the type library, the offical one, due to the end of support in .NET Core and .NET 5+, is to manually create a `.idl` and use the MIDL compiler to generate the `.tlb`. Another way is to continue to use .NET Framework just for regasm's capability to create a `.tlb` from a `.cs` file. Another one is to use the third party tool [dscom](https://github.com/dspace-group/dscom).

1. Download the `dscom32.exe` and the `dscom.exe` from the [dscom releases](https://github.com/dspace-group/dscom/releases).
2. Copy the dscom32.exe` and the `dscom.exe` to a `binaries/` folder on the `.sln` root.
3. Open your `.csproj` and add the following MSBuild command (For x86 only at the moment, there is a better way for x64)
```xml
<PropertyGroup>
  <WarningsAsErrors>$(WarningsAsErrors);DSCOM001</WarningsAsErrors>
</PropertyGroup>

<PropertyGroup>
  <SkipManualTlbExport>false</SkipManualTlbExport>
</PropertyGroup>

<Target Name="GenerateTypelib" AfterTargets="CoreBuild" Condition="'$(SkipManualTlbExport)' != 'true'">
  <Exec Command="$(SolutionDir)/binaries/dscom32.exe tlbexport $(TargetPath) /out:$(ProjectDir)$(TargetName).tlb" />
</Target>
```
4. Open an **elevated** command prompt and run `dscom32.exe tlbregister ProjectName.tlb`.
