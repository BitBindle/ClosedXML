**********************
ClosedXML Installation
**********************

====================
Installing ClosedXML
====================
The release workflow publishes ClosedXML to the GitHub Packages feed of BitBindle. The feed needs
authentication and a package source mapping. The `README <https://github.com/BitBindle/ClosedXML#install>`_
shows the ``nuget.config``. Then add the package with the .NET CLI.

.. code-block:: batch

   C:\source> dotnet add package ClosedXML --prerelease

==========================
Compatible implementations
==========================
ClosedXML targets .NET 10. It runs on .NET 10 and later, for example in Blazor.
ClosedXML doesn't work on .NET Framework, Mono or Unity.
