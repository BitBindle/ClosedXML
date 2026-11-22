#############################
Migration from 0.106 to 0.107
#############################

*****************
Target frameworks
*****************

ClosedXML now targets only ``net10.0``. This version does not have the ``netstandard2.0`` and
``netstandard2.1`` builds. .NET Framework, Mono and Unity applications can't use this version.
These applications can stay on version 0.106.

*****************
Equality methods
*****************

``XLColor.Equals(object)`` now returns false for null and for an object that is not an ``XLColor``.
Before, it threw an exception. The ``IEquatable<T>.Equals`` methods of a cell fill, of a differential
number format and of a differential protection now return false for null. Before, they threw
``NullReferenceException``.
