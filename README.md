# What is this?
This is an open source interpretation of what the MSVBVM60 library does. Such library is used in VB6 applications as a base-library, where most VB functions are implemented.

# Why?
Well, there's no need to do this, but since VB6 is getting outdated, an open source runtime library might add some years to VB6's lifespan.

# How we did this?
Since the source code of this library is not available, and reverse engineering the library is a no-go, we've used some tools to get the information we need.
For instance, IDA was used against a dummy VB6 .exe file where the target function was called.
API Monitor from ROHITAB was also used to check some runtime values and function calls from the original MSVBVM60.dll.
Some reverse-engineering sites such as the ones listed in http://sandsprite.com/vb-reversing/

# How to contribute
In order to contribute, please clone this repo, build it, and after you can run the example VB6 project, you can start adding new functionality.
To do so, try to add a test in the dummy VB6 project and check if that doesn't work against this library. Normally these kind of errors are related to missing exported functions, so, it's easy to check what the original library did (using the methods described before).
After the functionality is implemented, create a pull request.

# Compiling the sources
You'll need Visual Studio 2017 to compile the library. Just open the solution and build it (either Debug or Release).
An example VB6 project is located inside the Debug folder. In order to compile it, you'll need VB6 installed. Compile the .exe of that project inside the Debug folder, and run it from either opening it directly from Explorer, or from "Run" in Visual Studio (launch configuration will open the VB6 project, and you can add debug the library from there).

# Status of the library
This library IS NOT A FULL replacement of the original library. So, don't expect to be a drop-in replacement.
Currently, a small subset of the original library are implemented.
Most notable ones:
- Creation of ActiveX objects (CreateObject)
- Access to the VBA object (Shell, Iif, etc.)
- DLL function calls
- MsgBox
- InputBox
