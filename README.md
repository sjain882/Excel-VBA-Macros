Macros created by me, unless otherwise specified.

### Embed Image Links

Info:

- Automatically embeds images in the place of image links

- Can be re-run over cells to "fill in gaps" non-destructively

- Saves the image with the document, instead of links which can go down

- WARNING: No undo available! Duplicate your file before proceeding

- Original source: by kresty @ https://superuser.com/a/1404121

- Reworked by me:

    - Images are no longer stretched to fill cell - now preserves aspect ratio

    - Images are now properly resized to fit perfectly inside cell

    - Images are centered in the cell if smaller than the cell

    - Macro executes on currently selected cells, rather than having to manually specify range every time

    - Code cleaned up and commented

- Example before: `LFRD 2025 Bus Previews - B4 - With Hyperlink.xlsm`

- Example after: `LFRD 2025 Bus Previews.xlsm`

<details>
  <summary>🖼 Preview</summary>
  
![Before](https://github.com/sjain882/Excel-VBA-Macros/blob/main/.github/Previews/Embed%20Image%20Links/Before.png?raw=true)

![After](https://github.com/sjain882/Excel-VBA-Macros/blob/main/.github/Previews/Embed%20Image%20Links/After.png?raw=true)

</details>

### Google Active Cell

- Search the contents of currently selected cell with Google / Google Images / Flickr

- Created by Booksticle, unmodified in this repository (except adding PtrSafe to ShellExecute declaration for 64-bit operation)

- RunProgram from [here](https://shortcut.booksticle.com/showlist2.asp?parent=77924) ([archive](https://web.archive.org/web/20250712142935/https://shortcut.booksticle.com/showlist2.asp?parent=77924))

- Google from [here](https://shortcut.booksticle.com/showlist2.asp?parent=111219) ([archive](https://web.archive.org/web/20250712142956/https://shortcut.booksticle.com/showlist2.asp?parent=111219))

- Demo [here](https://www.youtube.com/watch?v=pAbmUdZgIMc) ([archive](https://web.archive.org/web/20201224062607/https://www.youtube.com/watch?v=pAbmUdZgIMc))

- Example included in `LFRD 2025 Bus Previews.xlsm`

<details>
  <summary>🛠 Setup Guide</summary>

1. Developer > Visual Basic

2. Modules > Module1 (in tree on left side)

3. Paste (Declarations).txt in Declarations section:

![Declarations](https://github.com/sjain882/Excel-VBA-Macros/blob/main/.github/Guides/Google%20Active%20Cell/Declarations.png?raw=true)

4. Tools > Macros

5. Under Macro Name type either `FlickrIt` / `GoogleIt` / `GoogleImagesIt` then click Create to create new macro section for the relevant function

6. New section with Sub skeleton will be created. Replace entire contents of this section with the contents of the relevant .txt containing the function's code

7. Save and close VBA Editor

8. Developer > Macros > Select macro you want to assign keyboard shortcut to > Options

9. Type a single key into shortcut box, e.g, G for CTRL+G. For CTRL+SHIFT+G, delete contents of box and type Shift+G (NOT Ctrl+Shift+G). Nothing else will work.

10. Click OK

</details>

<details>
  <summary>💼 Archived macros</summary>
None
</details>

***

# External

*(Macros & resources)*

[Booksticle](https://shortcut.booksticle.com) - Assorted macros, shortcuts, Excel guides