# Crystal Reports importing tool

The utility converts the Crystal Reports templates (.rpt-files) to the Stimulsoft Reports report templates format (.mrt-files). The tool is supplied as the C# source code only and requires referencing of some Crystal Reports runtime libraries in order to be built successfully in Visual Studio 2010, .NET Framework 4.0 or higher. Please download the archive from the link below, unzip it and open in the Visual Studio. The project will be built successfully, once all the required dll libraries are referenced and found in Visual Studio.

The project was created in a way that all the required assemblies would be automatically taken from the GAC (Global Assembly Cache). If Stimulsoft Reports dlls are not in the GAC, then the references to Stimulsoft.Base.dll  and  Stimulsoft.Report.dll must be added manually.

The Crystal Reports report templates’ file format is a proprietary format. Therefore, the tool requires some Crystal Reports special managed assemblies. The tool interacts with these assemblies via some special Crystal Reports interfaces for the special Visual Studio managed dlls.

These assemblies are not always installed in the system together with Crystal Reports, usually the additional and an official installation of these assemblies is required in order for them to work correctly with the import tool.

For example, for Crystal Reports 2013 the Support Pack (developer version for VS: Updates & Runtime) is required and needs to be installed first, and only after that the import tool will be built successfully.

The current Crystal Reports version requires the additional installation of the ‘SAP Crystal Reports runtime engine’ (32 bit or 64 bit). The automatic installer will copy the required assemblies to the GAC. But this installer must be downloaded separately, it is not a part of the standard Crystal Reports installation package.

The project uses the following Crystal Reports assemblies:
* CrystalDecisions.CrystalReports.Engine
* CrystalDecisions.ReportAppServer.DataDefModel
* CrystalDecisions.ReportAppServer.ReportDefModel
* CrystalDecisions.Shared
* CrystalDecisions.Web
* CrystalDecisions.Windows.Forms

These assemblies are not included with the tool. The packages will not work if they are just referenced and copied to the project without the proper installation by the Crystal Reports’ official installer first.

## Please find the explanation of the required installations

Operational system | Platform Target, CPU | Installation package requirements
------------------ | -------------------- | ---------------------------------
Windows x32 | Any CPU | ‘SAP Crystal Reports runtime engine 32 bit'.
Windows x64 | Any CPU | ‘SAP Crystal Reports runtime engine 64 bit'.
Windows x64 + runtime engine x32bit | X86 | not required
Windows x64 + runtime engine x32bit | Any CPU | ‘SAP Crystal Reports runtime engine 64 bit'.

## The above mentioned installers can be downloaded using the following links

http://www.crystalreports.com/crvs/confirm/
http://downloads.businessobjects.com/akdlm/cr4vs2010/CRforVS_redist_install_32bit_13_0_20.zip
http://downloads.businessobjects.com/akdlm/cr4vs2010/CRforVS_redist_install_64bit_13_0_20.zip

Please read more about the requirements of those additional installations in the official reply from the Crystal Reports:

https://archive.sap.com/discussions/thread/3675145

## Run Crystal Reports on client machine without install runtime package?

No, the only way to make your app work is to run one of the redist packages on the user’s PC. We don't support nor do we have a way to manually deploy the runtime. Too many Registry entries and registering of the dll's in order to do this manually.
