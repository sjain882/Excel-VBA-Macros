Macros created by me, unless otherwise specified.

# Macros

### Embed Image Links

Info:

- Automatically embeds images in the place of image links

- Can be re-run over cells to "fill in gaps" non-destructively

- Works with URLs not ending in image extension, e.g, discord image link with token query param

- If you add rows/columns around imported images, the embedded images move automatically

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

3. Paste (Declarations).vba in Declarations section:

![Declarations](https://github.com/sjain882/Excel-VBA-Macros/blob/main/.github/Guides/Google%20Active%20Cell/Declarations.png?raw=true)

4. Tools > Macros

5. Under Macro Name type either `FlickrIt` / `GoogleIt` / `GoogleImagesIt` then click Create to create new macro section for the relevant function

6. New section with Sub skeleton will be created. Replace entire contents of this section with the contents of the relevant .vba containing the function's code

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

# Other

### How to print spreadsheet as single-page zoomable PDF?

1. File > Print > Page Setup

2. Scaling > Fit To > 1 pages wide by 1 tall

3. OK

4. Select Portrait/Landscape as appropriate

5. Print

### How to print spreadsheet with working Hyperlinks?

Normal File > Print will not work - you need to use Export.

1. First, Print to PDF with above instructions to save print layout

2. Then, File > Export > Create PDF/XPS document

***

# Bonus

If you were browsing a lot of Flickr images then opened their direct links to paste into a sheet (with Embed Image Links), but want to go back and get the original full webpage link so you can credit the author, use this bookmarklet ([source](https://www.flickr.com/groups/96035807@N00/discuss/72157721920375928/72157721920408447)):

```javascript
javascript:const%20fid%20%3D%20(%2F%5Ehttps%3F%3A%5C%2F%5C%2F(farm%7Clive)%5Cd*%5C.static%5C.%3Fflickr%5C.com%5C%2F%5Cd%2B%5C%2F(%3F%3CimageId%3E%5Cd%2B)_%5B0-9a-f%5D%2B%5B%5E.%5D*%5C.%2Fiu).exec(document.URL)%3F.groups.imageId;%0Aif%20(fid)%20{%0A%20%20%20%20window.location%20%3D%20%60https%3A%2F%2Fflickr.com%2Fphoto.gne%3Fid%3D%24{fid}%60;%0A}%20else%20{%0A%20%20%20%20alert(%22Sorry%2C%20this%20is%20not%20a%20Flickr%20static%20image.%22);%0A}
```

***

# External

*(Macros & resources)*

[Booksticle](https://shortcut.booksticle.com) - Assorted macros, shortcuts, Excel guides