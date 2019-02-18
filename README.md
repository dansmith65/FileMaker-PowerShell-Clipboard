## Purpose

Convert FileMaker clipboard format both to and from text. Either input will result in clipboard being able to be pasted into either FileMaker or a text editor.

Why would you want to do this?

- Use a text editor find/replace within all steps from a script, which could then be pasted back into FileMaker.
- Quickly get the internal id of any FM element that can be copied.
- Extract an SVG icon from a layout object.
- Modify layout object styles that FileMaker interface doesn't allow you to (beware!).
- Save commonly used elements (like script steps) in a program like [PhraseExpander][], so you can quickly paste them without having to re-create them.
- Take the prior item to the next level with some user input and customize the clip before pasting it https://youtu.be/-DWA9i2eD3c .
- I could go on, but i'll stop there!



## Instructions

1. Download the script and put it in any folder you choose.
2. To run the script with the default options you can right-click on it and select **Run with PowerShell**. If you get an error about [Execution Policy][], you may need to modify the execution policy, or [temporarily bypass it][].
3. When the script runs without error and successfully detects/converts formats, it will immediately close. If there is an error, the window will stay open so you can view it.
4. I'd recommend setting up a hotkey to run it. I use <kbd>Alt</kbd> + <kbd>F2</kbd> defined in [PhraseExpander][], but there are many ways to setup a hotkey to run a program. Call it like this:
   ```
   powershell.exe -sta -file "C:\Path\To\Convert-FMClip.ps1"
   ```
5. [OPTIONAL] If you don't want the XML to be pretty printed, or if you prefer spaces over tabs, you can modify the [Set-Configuration.ps1](Set-Configuration.ps1) file, then run it.

If you use a Mac, this project isn't for you. You can Find a similar set of scripts written in AppleScript here: https://github.com/DanShockley/FmClipTools .



[Execution Policy]: https://docs.microsoft.com/en-ca/powershell/module/microsoft.powershell.core/about/about_execution_policies
[temporarily bypass it]: https://blogs.technet.microsoft.com/ken_brumfield/2014/01/19/simple-way-to-temporarily-bypass-powershell-execution-policy/
[PhraseExpander]: https://www.phraseexpander.com/