This archive contains the Windows 32 (ANSI) Type Library written by Joe Hart.

There is a test project included which will test the type library and make sure it is working properly.

To use the type library, simply place the jhwin32.tlb file in your WINDOWS\SYSTEM folder and run Visual Basic.  Load a project or create one, then choose Project, References from the menu.  Click the checkbox for the Windows API (ANSI) entry in the list.  If it isn't there, click brouse and find the jhwin32.tlb file, open that.

Once you have a reference to the library, you can then use any of the functions, constants or structures (user defined types) that are contained in the library, which are most of the API functions you would ever use, and quite a few you will never need.  You DO NOT have to copy information from the WIN32API.TXT file (which is wrong in many cases).  You can use the functions just like the API were a part of VB.

If you would like to know what exactly is in the TLB, you can use the object browser (F2 from VB) and select the WIN32.  It will list ALL of the Functions, Types and Constants that are contained, as well as give a brief description of what each of the functions is for.

When you compile your program, VB will look at the type library and link ONLY THE FUNCTIONS YOU USE, so you don't have to worry about your program being so large.  You basically have the entire API at your disposal with NO OVERHEAD.  When you distribute your program, you DO NOT have to give the user the type library.  Only VB needs it.  You will have to include it if you want others to be able to compile your programs (meaning you give them the source code).  Although it can be a mean trick to give people source code but keep them from compiling your program by not giving them the type library.  If they don't have the library they can run the EXE, but they can't run, or compile the source code.

Why use the type library?  Well, the main reason is if you use the API, you rarely have to worry about Declare statements.  You never have to hunt for what the value of one of the constants is if your reading the Windows SDK and all it lists is the names of the constants. Just use the same name you see and most likely that's the name in the library.  If not it's because it would interfere with VB (like the structure POINT is called POINTL because POINT is a reserved word in VB).


This type library was written by Joe Hart. based on ideas presented in Bruce McKinney's Hardcode Visual Basic.