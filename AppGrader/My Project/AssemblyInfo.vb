Imports System.Resources

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("AppGrader")> 
<Assembly: AssemblyDescription("There is a lot of administration needed when grading assignments submitted through Blackboard, especially when each submission is zipped. Errors due to excessively long file names can result. This application trims each zip filename leaving only the user ID, and then unzips these files placing the contained files in their own subdirectory. It was designed to handle Programming Homework, but it will handle any compressed file submission.")> 
<Assembly: AssemblyCompany("University of South Carolina")> 
<Assembly: AssemblyProduct("Application Grader")> 
<Assembly: AssemblyCopyright("Copyright © John Gerdes 2015")> 
<Assembly: AssemblyTrademark("")> 

<Assembly: ComVisible(False)> 

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("65ddfc49-fb78-496b-8759-7eb5b1fc8ea9")> 

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

<Assembly: NeutralResourcesLanguageAttribute("en-US")> 