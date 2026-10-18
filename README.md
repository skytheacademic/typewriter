# Typewriter

**A Microsoft Word add-in that makes you write forward.** Turn it on and you can only add text to the end of your document: no deleting, no going back to fix things, no fiddling. Set a timer, get the draft down, and edit later.

The worst thing to do while drafting is edit. Typewriter takes that option away for as long as you choose.

## What it does

While a session is running, in that document:

- **Deleting is blocked:** Backspace, Delete, and the word-at-a-time versions (Ctrl+Backspace, Ctrl+Delete, and their Mac equivalents).
- **Cut, Undo, and Find and Replace are blocked**, from the keyboard, the ribbon, and the right-click menu.
- **The cursor stays at the end.** Clicking into old text, using the arrow keys, or selecting text snaps the cursor back to the end, so you can't type over what you wrote. You can still scroll up and reread.
- **Distractions are paused:** AutoCorrect replacements, spelling and grammar squiggles, drag-and-drop, and Overtype mode. Your settings come back when the session ends.

You choose how long a session lasts: 5 to 90 minutes, or no limit. **A timed session can't be stopped early.** When the time is up, editing unlocks and Typewriter tells you how many words you wrote. A session with no limit can be stopped at any time, after a confirmation.

If you really need out, closing the document always ends the session.

## Install

Download **Typewriter.dotm** from the [latest release](https://github.com/skytheacademic/typewriter/releases/latest). It's one file for both Windows and Mac.

You need desktop Word (Microsoft 365, or Word 2016 or later). It does not work in Word on the web, on iPad, or on phones.

### Windows

1. Close Word.
2. **Unblock the file.** Windows blocks macros in downloaded files. Right-click `Typewriter.dotm` → **Properties** → tick **Unblock** at the bottom → **OK**.
3. Press **Win+R**, paste `%APPDATA%\Microsoft\Word\STARTUP`, and press Enter.
4. Copy `Typewriter.dotm` into that folder.
5. Open Word. There's now a **Typewriter** tab on the ribbon.

### Mac (beta)

1. Quit Word.
2. In Finder, choose **Go → Go to Folder…** (Shift+Cmd+G), paste the path below, and press Return:
   ```
   ~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word
   ```
   If it isn't found, open Word, go to **Word → Settings → File Locations → Startup** to see the right folder, then quit Word again.
3. Copy `Typewriter.dotm` into that folder.
4. Open Word. If it asks about macros in Typewriter.dotm, choose **Enable Macros**.

If the Typewriter tab doesn't appear:
- Set **Word → Settings → Security & Privacy → Macro Security** to *Disable all macros with notification*, then restart Word.
- If the file came from a download, AirDrop or email, run this in Terminal and restart Word:
  ```bash
  xattr -c ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Startup/Word/Typewriter.dotm
  ```

## Use

On the **Typewriter** tab:

- **Session**: pick a length (5–90 minutes) or *No limit*.
- **Start Writing**: locks the current document. The tab and the window title show the time left.
- **Diagnostics**: checks that everything works on your computer and shows a pass/fail report. Useful for bug reports.
- **Check for Updates**: asks GitHub whether a newer version exists (Windows; Mac shows the releases link).

Keyboard shortcut: **Ctrl+Alt+T** on Windows. This replaces Word's shortcut for the ™ symbol.

## Updating

Typewriter doesn't update itself. To update, quit Word, download the new `Typewriter.dotm` from [Releases](https://github.com/skytheacademic/typewriter/releases), replace the old file in the same folder (unblocking it first on Windows), and reopen Word. Your chosen session length is kept.

To hear about new versions, click **Watch → Custom → Releases** at the top of this page.

## Uninstall

Quit Word and delete `Typewriter.dotm` from the Startup folder above.

## Known issues

- **Mac support is beta, and its known problems are all keyboard shortcuts.** Tested on Word 16.114 for Mac, blocking the delete key, keeping the cursor at the end, the session timer and pausing spellcheck all work. Some shortcuts may not work, or may not be blocked:
  - the start shortcut (Control+Option+T). Use the **Start Writing** button instead.
  - Option+delete and Cmd+delete, and their fn+delete versions.

  Running **Diagnostics** shows which shortcuts work on your Mac.
- **The Windows keystroke tests ran on one machine.** Please open an issue with your Diagnostics report if something gets through.
- **If Word crashes during a session**, AutoCorrect, spellcheck-as-you-type and drag-and-drop stay switched off. Turn them back on in Word's options.
- **It's a commitment device, not a lock.** You can always close the document, or remove the add-in.

## Privacy

Typewriter collects nothing. The only network request it ever makes is the **Check for Updates** button, which asks `api.github.com` for the latest release, and only when you click it.

## How it works

Typewriter is a Word template with VBA macros (`.dotm`) that Word loads from its Startup folder.

- **Blocking commands:** macros named after Word's built-in commands (`EditCut`, `EditUndo`, `EditClear`, …) replace those commands, so the ribbon, menus and shortcuts all go through Typewriter.
- **Blocking keys:** the delete keys are bound to Typewriter's own macros. When no session is running, each does what the key normally does.
- **Keeping the cursor at the end:** a Word event moves the cursor back to the end whenever the selection changes.
- **The timer:** `Application.OnTime` ticks once a second. If it's ever missed, the session still ends the next time you type, click, or press the button.
- **The ribbon tab:** a `customUI14.xml` part inside the `.dotm`.

## Building from source

You need Windows, desktop Word, and PowerShell 7 (`pwsh`).

1. In Word, switch on **File → Options → Trust Center → Trust Center Settings → Macro Settings → Trust access to the VBA project object model**. The build uses it to put the code into the template. Switch it off again afterwards.
2. Close Word, then from the repo folder:
   ```bash
   pwsh build/build.ps1 -Install
   ```
   This creates `dist/Typewriter.dotm`, and with `-Install` copies it into Word's Startup folder.
3. Run the automated checks (they open and close Word):
   ```bash
   pwsh tests/test.ps1
   ```

Source files:

| File | What it is |
|---|---|
| `src/Typewriter.bas` | All the logic: sessions, timer, key handling, ribbon callbacks, diagnostics, update check |
| `src/TypewriterEvents.cls` | Listens for Word events (selection changes, documents closing) |
| `src/customUI14.xml` | The Typewriter ribbon tab |
| `build/build.ps1` | Builds the `.dotm` from `src/` |
| `build/add-ribbon.ps1` | Adds the ribbon tab to the built file |
| `tests/test.ps1` | Automated checks against an installed copy |

To release a new version, bump `TYPEWRITER_VERSION` in `src/Typewriter.bas`, build, and attach `dist/Typewriter.dotm` to a GitHub release tagged `vX.Y.Z`. **Check for Updates** compares against that tag.

## License

[MIT](LICENSE) © Sky Kunkel
