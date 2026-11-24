**********************
ClosedXML Installation
**********************

====================
Installing ClosedXML
====================
The ``BitBindle.ClosedXML`` package is on nuget.org. The package ID differs from the ID of the original
project. The assembly name and the namespaces are the same, so a project cannot use both packages.
Add the package with the .NET CLI.

.. code-block:: batch

   C:\source> dotnet add package BitBindle.ClosedXML

==========================
Compatible implementations
==========================
ClosedXML targets .NET 10. It runs on .NET 10 and later, for example in Blazor.
ClosedXML doesn't work on .NET Framework, Mono or Unity.
