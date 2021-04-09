# Define styles for a new Word&nbsp;template: first example

To make a new Word template means defining many styles. The easiest way I've found is to write a description of the styles, then run a macro that defines styles like the description.

For example, suppose you want a template that matches the format of an ANSI standard, "[Scientific and Technical Reports&mdash;Preparation, Presentation, and Preservation](https://www.niso.org/publications/z39.18-2005-r2010)." Measure its margins, compare the font sizes to some samples, and write a description. Use the format shown below, with a style name at the beginning of the line and commas between the specifications.

Other specifications can be included from the [list of available specifications](https://www.mechanicaledit.com/macro_specs). Check back later for [more sample style descriptions](https://www.mechanicaledit.com/style_descriptions) for other types of documents.

## Apply the style descriptions

To apply style descriptions, open a new Word document, set the style defaults, paste the style descriptions, paste and run the style macro, and save the Word document.

####  Open a document without styles

1. Type **winword /a /w** in the Windows taskbar and press **Enter**. \
<span style='font-size:small; color:darkgray;'>&#128712; The /a switch opens Word without opening your Normal template, which might have custom styles. The /w switch opens a new blank document. For more info see [Command-line switches for Microsoft Office products](https://support.microsoft.com/en-us/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6).</span>

#### Set the style defaults

1. In Word, click the **Design** menu and click **Fonts**.
1. Select the theme fonts. For these sample styles, click **Arial**.
1. In the search box in the menu bar, type **styles** and press **Enter**.
1. Click **Manage Styles**, the third button in the Styles pane.
1. Click the **Set Defaults** tab in the Manage Styles window. 
    1. Select a font size. For these sample styles, select **10**. 
    1. For the paragraph spacing after, select **0 pt**. 
    1. Select the line spacing. For these sample styles, leave **Multiple** and type **1.04**.
    1. Click **OK**. 

#### Add the style descriptions and macro

1. Copy the style descriptions (see above).
1. Right-click the Word document and select the paste option **Keep Text Only**.
1. Copy the macro (see below).
1. In the search box in the menu bar, type **visual basic** and press **Enter**.
    1. In the Visual Basic window, click **Insert** and **Module**.
    1. Click **Edit** and **Paste**.
    1. Click **File** and **Close and Return to Microsoft Word**.
1. In Word, click the **View** menu and click **Macros**.
1. Select the macro **sctApplySpecs** and click **Run**.

#### Save the file

1. Click **File** and **Save As**.
1. Click **Browse**.
1. Select a folder.
1. Type a file name. For the sample styles, type **Sample standard styles**.
1. Select a file type. \
&bull; To start a document, leave **Word Document (\*.docx)** as the file type. \
&bull; To start a template, leave **Word Template (\*.dotx)**. \
&bull; To make further style changes with the macro, select **Word Macro-Enabled Document (\*.docm)**.
1. Click **Save.** \
Click **Yes** in response to "...Continue saving as a macro-free document?"

---

### Legal

Copyright (C) 2021 Jay Martin. 

**Permission is granted** to copy, distribute and/or modify this document
under the terms of the [GNU Free Documentation License, Version 1.3](https://www.gnu.org/licenses/fdl-1.3.txt)
or any later version published by the Free Software Foundation; 
with no Invariant Sections, no Front-Cover Texts, and no Back-Cover Texts.
A copy of the license is included in the section entitled "[GNU Free Documentation License](fdl-1.3.md)."

Microsoft Windows is a trademark of Microsoft. All other trademarks are the property of their respective owners. 

<!--- --->
