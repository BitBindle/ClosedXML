############
Missing font
############

ClosedXML measures text to calculate column widths and row heights, for
example in ``IXLColumn.AdjustToContents``. To measure text, ClosedXML needs
the metrics of the font. If the font of the text is not on the computer,
ClosedXML uses a different font. The calculated widths and heights can then
be different from Excel.

This problem occurs most frequently on Linux, in containers and in cloud
functions (for example AWS Lambda and Azure Functions). These computers
usually do not have the Windows fonts, for example *Arial* or
*Times New Roman*. For *Calibri*, the default font of a workbook, ClosedXML
has an embedded font with almost the same widths.

.. note::
   For more information about the engine and its API, see the
   :doc:`../features/graphic-engine` page.

************************
How ClosedXML gets fonts
************************

The graphic engine measures text. The default engine is
``DefaultGraphicEngine``. It uses the *SixLabors.Fonts* library to read the
font metrics.

``DefaultGraphicEngine`` tries these fonts in this sequence:

#. The font of the text, for example the font of the cell.
#. The fallback font. The default fallback font is *Microsoft Sans Serif*.
   Excel uses the same fallback font on Windows.
#. The embedded font. This font is in the ClosedXML assembly.

The embedded font is a reduced version of *Carlito*. It contains only the
metrics of the glyphs. It has the same character widths as *Calibri*. With
the embedded font, ClosedXML uses the ascender and the descender of
*Calibri*. The row heights are then the same as with *Calibri*.

ClosedXML always has the embedded font, so it does not throw an exception
for a missing font. Versions before 0.102 threw this exception:

  Unable to find font *font name* or fallback font *fallback font name*.

Calibri and Carlito
===================

*Carlito* is metric-compatible with *Calibri*. Each character has the same
width in the two fonts. This is correct for large font sizes. At sizes up to
19 points, *Calibri* uses bitmaps. The widths of these bitmaps are a little
different.

*Carlito* is a little wider than these bitmaps, for example the letter M and
the digits. ClosedXML calculates the column widths with the
*Carlito* metrics, so the columns are a little wider than the text in Excel.
If a column is too narrow, Excel cuts the text or shows ``####``.

The full *Carlito* font has its own ascender and descender. With the full
font, the row heights can be a little different from *Calibri*.

.. image:: img/Calibri-vs-Carlito-11pt.png
  :alt: Comparison of Calibri and Carlito font rendered at 11 points

.. image:: img/Calibri-vs-Carlito-20pt.png
  :alt: Comparison of Calibri and Carlito font rendered at 20 points

**************************
When the widths are wrong
**************************

A missing *Calibri* font usually causes no problem. ClosedXML uses the
embedded font. The widths are then almost the same as in Excel. But the
fallback font comes before the embedded font. If the computer has
*Microsoft Sans Serif* but not *Calibri*, ClosedXML uses
*Microsoft Sans Serif* for *Calibri* text.

The widths can also be wrong in these conditions:

* The workbook uses a font other than *Calibri*, for example *Arial* or
  *Times New Roman*, and the computer does not have this font. ClosedXML
  measures the text with the fallback font or the embedded font. These fonts
  have different metrics.
* The font that ClosedXML selects does not have all the characters of the
  text. ClosedXML does not use a different font for each missing character. For
  example, the embedded font has approximately 2500 glyphs for the Latin,
  Greek and Cyrillic alphabets. It does not have glyphs for other writing
  systems, for example Chinese, Japanese, Korean or Arabic.

**********************************************
Where SixLabors.Fonts looks for system fonts
**********************************************

System fonts are the fonts on the computer. *SixLabors.Fonts* looks for
system fonts in these folders (see the
`source <https://github.com/SixLabors/Fonts/blob/main/src/SixLabors.Fonts/SystemFontCollection.cs>`_):

* Windows

  * ``%SYSTEMROOT%\Fonts``
  * ``%APPDATA%\Microsoft\Windows\Fonts``
  * ``%LOCALAPPDATA%\Microsoft\Windows\Fonts``

* Linux

  * ``%HOME%/.fonts/``
  * ``/usr/local/share/fonts/``
  * ``/usr/share/fonts/``

* macOS

  * ``%HOME%/Library/Fonts/``
  * ``/Library/Fonts/``
  * ``/System/Library/Fonts/``
  * ``/Network/Library/Fonts/``

* Android

  * ``/system/fonts/``

*********
Solutions
*********

Install the fonts of the workbook
=================================

Install the fonts that the workbook uses on each computer that runs your
application. ClosedXML then uses the correct metrics.

* *Calibri*, *Arial* and other Microsoft fonts need a license.
* On some computers, you cannot install fonts, for example in Azure
  Functions.

Bundle the fonts with the application
=====================================

Put the font files in your application, for example as embedded resources.
Make sure that the license of each font lets you distribute it. Then give the
font streams to ``DefaultGraphicEngine``. The first stream is the fallback
font. The other streams are optional. Add the fonts that your workbooks use.

ClosedXML finds a font by its family name. A stream helps only the text with
a font name that is the same as the family name in the stream. For example,
a *Liberation Sans* stream does not help *Arial* text.

.. code-block:: csharp

   using var fallbackFontStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Namespace.Carlito.ttf");
   using var arialFontStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Namespace.Arial.ttf");
   LoadOptions.DefaultGraphicEngine = DefaultGraphicEngine.CreateWithFontsAndSystemFonts(fallbackFontStream, arialFontStream);

``DefaultGraphicEngine.CreateOnlyWithFonts`` has the same parameters, but it
does not load the system fonts. If the application cannot use a file system,
for example in client-side Blazor, use this method. It also makes the first
measurement of text faster.

With ``CreateWithFontsAndSystemFonts`` and ``CreateOnlyWithFonts``, the
fallback font is always available. ClosedXML then never uses the embedded
font. If the workbook uses *Calibri*, give a *Calibri* stream or use
*Carlito* as the fallback font. ``CreateWithFontsAndSystemFonts`` also uses
*Calibri* if the computer has it.

Set a different fallback font
=============================

If an almost correct width is enough, set a system font as the fallback
font. This font replaces *Microsoft Sans Serif* in the sequence. If the
computer does not have *Calibri*, ClosedXML uses this font for *Calibri* text,
not the embedded font. To get a list of the system fonts, use this code:

.. code-block:: csharp

   foreach (var fontFamily in SixLabors.Fonts.SystemFonts.Collection.Families)
      Console.WriteLine(fontFamily.Name);

Select a font that has character widths similar to the font of the workbook.
A font that is a little wider is better. The columns are then a little wider
than the text in Excel. Set the font name, not the path of the font file:

.. code-block:: csharp

   LoadOptions.DefaultGraphicEngine = new DefaultGraphicEngine("DejaVu Sans");

For workbooks that use *Calibri*, the full
`Carlito font <https://fontlibrary.org/en/font/carlito>`_ is a good fallback
font. It has the same widths as *Calibri*. If the computer has
*Microsoft Sans Serif* but not *Calibri*, install *Carlito* on the computer.
Then set it as the fallback font:

.. code-block:: csharp

   LoadOptions.DefaultGraphicEngine = new DefaultGraphicEngine("Carlito");

Set the column widths and row heights
=====================================

Set ``IXLColumn.Width`` and ``IXLRow.Height`` directly. ClosedXML then does
not calculate these values from the text. The unit of the row height is
points. The unit of the column width is the number of characters.

The maximum digit width is the width in pixels of the widest digit (0 to 9) of
the default font. For a width of 1 or more, the pixel width is the number of
characters times the maximum digit width, plus a padding. For example, with
*Calibri* 11 pt at 96 DPI, the maximum digit width is 7 pixels and the padding
is 5 pixels. A column with width 12 is then 89 pixels.

Implement a graphic engine
==========================

Write a class that implements ``IXLGraphicEngine``. In
``IXLGraphicEngine.GetPictureInfo``, call
``DefaultGraphicEngine.Instance.Value.GetPictureInfo``. In the methods that
measure text, calculate the values from the font size
(``IXLFontBase.FontSize``) or use fixed values. The results are usually less
accurate than the results of ``DefaultGraphicEngine``.
