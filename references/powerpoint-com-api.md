# PowerPoint attach-only COM API

The adapter attaches to existing desktop document windows. It never creates an Application, opens files, uses ActivePresentation, activates document windows, saves, closes or quits user content. The live test alone creates and closes owned unsaved presentations; it never calls Application.Quit because PowerPoint can share one singleton process with the user.

## Identity and connection evidence

list enumerates PPTFrameClass windows via EnumWindows, reporting HWND/PID. probe/read require explicit positive --pid and --hwnd. Captions are descriptive output only.

Two read-only connection paths are supported:

1. NativeOM: AccessibleObjectFromWindow with OBJID_NATIVEOM and IID_IDispatch on a paneClassDC descendant. Microsoft specifies a DocumentWindow result. Validate the pane PID and GetAncestor(GA_ROOT) against the requested top-level HWND, then use DocumentWindow.Presentation.
2. ROT fallback: GetActiveObject(PowerPoint.Application) obtains only a registered running object. Enumerate Presentations and each Presentation.Windows within a combined 100-window budget. Use the installed type library's early-bound DocumentWindow.HWND vtable getter. GetAncestor(GA_ROOT) and GetWindowThreadProcessId map this document HWND to the requested top-level HWND/PID. Require exactly one match. No document selection is inferred from ROT order, window title, active state, or a file path.

Actual local evidence: NativeOM's documented paneClassDC class was absent from new owned PowerPoint frames (only Mso/NetUI/MDI children). Late-bound DocumentWindow.HWND returned DISP_E_MEMBERNOTFOUND. Early binding through comtypes GetBestInterface succeeded and returned an mdiClass child HWND. The installed type-library getter plus Win32 root/PID validation is therefore the working local path. This hidden getter is version dependent; missing or ambiguous identity fails closed. No arbitrary vtable offsets are hardcoded.

Unsaved decks are supported: name, exact document HWND, root HWND and PID establish identity. full_name is the COM FullName value; for unsaved decks it may be just a name, not a disk path. --document compares that value only. All reads rebind after retrieval; these checks detect some races and are not an atomic snapshot or a concurrency lock.

Official sources researched:

- [AccessibleObjectFromWindow](https://learn.microsoft.com/en-us/windows/win32/api/oleacc/nf-oleacc-accessibleobjectfromwindow): paneClassDC returns DocumentWindow with OBJID_NATIVEOM.
- [GetActiveObject](https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-getactiveobject): retrieves an OLE-registered running object.
- [GetAncestor](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getancestor): GA_ROOT follows the parent chain to the root window.
- [DocumentWindow](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindow): Presentation property associates a document window with its presentation. HWND is a hidden installed-type-library member, not listed on this VBA page; its operational evidence comes from the independent live test.
- [DocumentWindows](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindows): open document-window collection excludes slide-show windows.
- [Slide.SlideID](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.slideid): identifier remains stable when slides are reordered.
- [Shape.Id](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.id): shape identifier, scoped with slide ID in evidence.
- [TextRange.Characters](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange.characters): bounded range retrieval after Length is checked.
- [Slide.NotesPage](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.notespage): notes are accessed only with --notes.
- [Table.Cell](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.table.cell): table cell access only with --tables.

## CLI and output

Run using the repository COM environment:

    .venv/Scripts/python.exe scripts/powerpoint_com.py list
    .venv/Scripts/python.exe scripts/powerpoint_com.py probe --pid PID --hwnd HWND
    .venv/Scripts/python.exe scripts/powerpoint_com.py read --pid PID --hwnd HWND --offset 1 --limit 2 --shape-offset 1 --shape-limit 30 --notes --tables --output tmp/ppt-read.json

All commands emit schema_version=1, status, operation, target, elapsed_seconds, limits and data. Errors go to stderr and return nonzero. Successful --output uses a temporary file in the same directory plus os.replace. Failures preserve any previous output; consumers must check exit status and target to avoid stale results. The 9-second parent watchdog kills only its own worker, never PowerPoint. Direct imported API functions have no watchdog and must be supervised by their caller.

Slide pages: offset >=1, limit 1..50 (default 5). Shape pages: shape-offset >=1, shape-limit 1..500 (default 50), applied independently to each selected slide and optional notes page. next_offset cursors indicate remaining slide/shape items. Total selected shapes including notes <=500. Tables <=500 cells total. Native text length <=32,000 per COM retrieval and <=100,000 total (UTF-16 accounting); --text-budget can reduce this. Budget overflow refuses the complete request before publishing. Whole Text retrieval occurs only on a bounded Characters range after native Length validation; empty text avoids Characters(1,0).

Slide ID and shape ID plus index/name provide native evidence. Text includes SHA-256 of UTF-8 and native length. A table cell is identified by slide ID + table shape ID + row/column; cell Shape.Id is unsupported on the tested build and is not used. IDs identify objects within a presentation and do not provide a global persistent document identity.

Groups/chart/OLE/SmartArt child content is explicitly omitted; no recursive traversal, link updating, embedded-object activation, OCR, chart series extraction, notes masters, themes or visual fidelity checks. Protected View, hidden presentations, slideshow-only frames, modal/busy Office, unregistered ROT instances, multiple process registrations or missing hidden HWND getter may not attach. Such cases report errors, not guessed targets.

## Validation

Use .venv/Scripts/python.exe -m unittest discover -s tests -p test_powerpoint_com.py -v and .venv/Scripts/python.exe tests/live_powerpoint.py --run.

Independent live validation succeeded with another owned presentation active, exercising exact HWND binding, empty deck, Chinese text, tables, notes, slide/shape IDs, slide pagination, wrong-PID refusal, text-budget failure and unchanged unsaved status. Supervised read CLI: 0.518 seconds wall time, 0.308 seconds worker time. Existing user window was recorded and preserved; cleanup touched only returned owned presentation objects and never quit PowerPoint. Initial one-time comtypes type-library generation took about 6.5 seconds; cold/busy Office can hit the 9-second watchdog and must fail safely rather than promise every request will succeed.
