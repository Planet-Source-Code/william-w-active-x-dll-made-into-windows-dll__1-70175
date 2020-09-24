read the original article http://www.windowsdevcenter.com/pub/a/windows/2005/04/26/create_dll.html?page=1
and also Joe Priestley www.jsware.net




Ever Wanted to make a Windows Dll but stuck with VB Active X Dlls?

Well, Here is your answer.


Compile the linkhelper as Link.Exe
go to your VB98 Directory (usually C:\Program Files\Microsoft Visual Studio\VB98\)


rename the original Link.exe to LinkLnk.exe and Copy our link Helper (Link.exe) 
to the VB98 Directory

This will intercept the calls to LinkLnk.exe and determine if its a dll with a .def file

If so it will ask you if you want to make a windows dll, if not or if there isn't a .def file

or the file isn't an active x dll it will just pass the arguments to Linklnk.exe and it will 

be as if nothing was changed except now with every project compiled there will also be a log

File called lnklog.txt which simply states the command arguments passed to Liknlnk.exe.

If you choose to make a windows dll the link helper will add /def (along with your .def file)

to the arguments and then Linklnk.exe will make a Windows Dll

optionally compile Make Def Addin to your VB98 Directory (usually C:\Program Files\Microsoft Visual Studio\VB98\)
to make your .def files for you & a declare helper text file


--------------------------------------
Notes:
--------------------------------------

  The compiler, C2.exe, can also be hooked into. The command lines
are roughly equivalent to the VC++ project settings tabs for C/C++ 
and linking. Presumably the compiler command line could be edited for finer
control of compiler options. The linker command line could also be
edited further to add LIB linkage, though it's not clear to me that that
would serve any purpose, as the C/C++ LIB linking seems to be what
VB does with a typelib, but not interchangeable. Nevertheless, an
inspection of the linker command line shows that a normal VB compile
is linking to the VB6 LIB.

  About strings:

   If you want to present a typical ANSI function, taking parameters in as
  ANSI and sending them back as ANSI, then it's necessary to
  use StrConv(s, vbUnicode) on the incoming strings and StrConv(s, vbFromUnicode) 
  on the string values being returned. 

    The reason for that is that VB is normally using unicode strings,
  while a VB programmer sees ANSI strings. By default, VB converts
  a string to unicode for its own internal operations, then converts
  it back to ANSI for DLL calls. With VB DLL exported functions, VB
  doesn't do its normal behind-the-scenes conversion with exported 
  functions. So a string coming in as ANSI must be converted for VB.
  Then when it's returned it must be converted back to ANSI "by hand".
  The sample code included here shows how that works.

  Calling a VB DLL from non-VB code:

   When a VB DLL is loaded, VB doesn't know to initialize the
runtime. If you want to call your DLL from C++ (or another
language where the VB  runtime won't be loaded) it requires
an extra step.

   There's a sample at vbadvance.com.  If you download the
vbAdvance package, what you get is apparently a VS add-in
program. That's not needed, but if you unpack the Inno
installer (use innounp or Universal Extractor) you'll find a folder 
with sample code.

   If you look in the \Exports\Non-VB Caller folder  you'll see a
sample project. Note that the sample includes (in addition to
the files in that folder) the two files in the Shared folder: 
MRuntimeInit.bas and CRuntimeInit.cls.

  As can be seen from the sample project, to call your DLL from 
C++, etc., you need to include those two files in your project, 
reference vbadvance.tlb, and call RuntimeInitialize in DLLMain.

  If you look at the help file from vbAdvance you'll find
further explanation: VB normally builds in code to load
the runtime, but doesn't know to do that when DLL exports
are created. So for non-VB code the extra functions
in MRuntimeInit.bas are added. They load the cRuntimeInit
class as a COM object, thereby getting the needed 
DLL initialization. 
  The vbadvance.tlb typelib is needed for the functions
in that call.

  Note: I haven't tested the self-initializing code option,
but there is a test sample in the vbAdvance download.

______________________________________________





