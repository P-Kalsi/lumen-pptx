# Lumen PPTX

![Lumen PPTX](images/icon.png)

Read PowerPoint files directly in a VS Code editor tab. Includes continuous slide
scrolling, a thumbnail sidebar, selectable text, search, slide navigation, zoom,
and speaker notes. Inspired by the viewing experience of tomoki1207's vscode-pdf;
this is an independent extension.

## Use

Open any `.pptx` and it loads in the preview automatically (this extension registers
as the default `.pptx` editor). To open it as raw XML instead, use **Reopen Editor
With… → Text Editor**.

Use the toolbar to show thumbnails, search, jump to a slide, zoom, or show notes.
Keyboard shortcuts: Cmd/Ctrl+F to search; left/right or Page Up/Page Down to change
slides; Home/End for the first/last slide; Cmd/Ctrl +/− to zoom and Cmd/Ctrl+0 to fit.
Pinch on a trackpad (or Cmd/Ctrl + scroll) zooms continuously. Text can be selected
and copied. Changes to the source file reload the viewer; the reload button rereads
it manually.

### File-explorer icon

The icon shown beside a file in the explorer comes from your active **File Icon
Theme**, not from any extension — VS Code has no API to set a per-extension file icon.
The built-in **Seti** theme (the default) ships `pdf`, Word and Excel icons but has no
PowerPoint mapping, so `.pptx` shows the generic file icon. To get a PowerPoint icon,
switch to an icon theme that includes one (e.g. **vscode-icons** or **Material Icon
Theme**) via *Preferences: File Icon Theme*.

## Files and storage

The extension reads the PPTX and sends its bytes to the webview in memory. The
renderer parses the Office XML and draws HTML/SVG. It does not create PDFs, images,
converted copies, a preview cache, or a separate preview tab. Closing the editor
releases its renderer. VS Code may retain the slide number and viewing preferences
as editor state. The source PPTX is read-only. It does not require LibreOffice,
PowerPoint, or vscode-pdf. No presentation data is uploaded; remote media is blocked.

## Rendering scope

Uses [@aiden0z/pptx-renderer](https://github.com/aiden0z/pptx-renderer) 1.4.0.
Supports text, images, shapes, tables and the renderer's chart/diagram features.
Legacy EMF images are rendered by the renderer; WMF images (which it does not
decode) are drawn from the PPTX with
[wmf-emf-renderer](https://www.npmjs.com/package/wmf-emf-renderer). Exact
PowerPoint fidelity is not guaranteed — installed/embedded fonts influence
appearance, text boxes that overlap tables may lay out imperfectly, and this is a
static viewer so animations/transitions are not played. Some SmartArt and embedded
objects may not render fully. All original presentation data remains in the PPTX.

## Development

`npm ci`, `npm run build`, `npm test`, `npm run package`.

`PPTX_DIR=/path/to/decks node test/browser.js` renders every `.pptx` in that folder
with installed Chrome (defaults to the parent folder).
`node test/vscode.js` launches an isolated VS Code profile and tests the real
webview. Test screenshots go to the system temporary directory.

Bundled dependencies and their notices are in `licenses/`.

## License

MIT © P-Kalsi. Not affiliated with or endorsed by Microsoft; PowerPoint is a
trademark of Microsoft Corporation.
