# TabSanity Visual Studio Extension

This package causes the Visual Studio text editor to treat tabs-as-spaces as
if they were actually tabs. That is, the backspace and delete keys, arrow key
navigation will not allow the caret to land within the spaces that form a tab.

## Installing

This package works with Visual Studio 2015. The easiest way to install
the package is with Visual Studio's built-in extension manager. Go to
`Tools | Extensions and Updates... | Online | Visual Studio Gallery` and search
for TabSanity. You can also find it on the [Visual Studio Gallery website](http://visualstudiogallery.msdn.microsoft.com/c8bccfe2-650c-4b42-bc5c-845e21f96328).

For Visual Studio 2013, please download and install [TabSanity.vsix](https://github.com/jedmao/tabsanity-vs/raw/master/TabSanity.vs2013/TabSanity.vsix). Thanks @FlipB!

## Building

1. Install the [Visual Studio SDK](http://msdn.microsoft.com/en-us/library/vstudio/bb166441(v=vs.110).aspx).
2. Open the solution file `TabSanity.sln` in [Visual Studio](http://www.microsoft.com/visualstudio/) 2015.
3. Look in `TabSanity/Bin/(Debug|Release)/TabSanity.vsix` and double-click to install.

## Dependencies

Though not "technically" required, [EditorConfig](http://visualstudiogallery.msdn.microsoft.com/c8bccfe2-650c-4b42-bc5c-845e21f96328)
has been added as a required dependency to this extension. This is to ensure
any `.editorconfig` files are picked up and applied before assuming the
document's tab settings. TabSanity, though completely decoupled from
EditorConfig, is designed to listen for changes in the text editor options
and adjust accordingly, just as EditorConfig does.

## Reporting problems

At this time, feel free to contact Jed Mao directly. See [humans.txt](https://github.com/jedmao/tabsanity-vs/blob/master/humans.txt)
file for contact information.
