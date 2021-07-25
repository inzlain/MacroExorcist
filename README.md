# Macro Exorcist
Macro Exorcist is a proof of concept tool to deliberally set an invalid VBA project version value in Word documents to ensure that compiled VBA code (pcode) is always recompiled when a document is opened. This is potentially useful for organisations reliant on documents with signed macros as it is possible to ensure that digitally signed documents will always recompile VBA code when opened rather than risk running the compiled and unsigned pcode.

## Background
Macros in Ofice documents are stored in the document as both source code and compiled code. The source code and the compiled code do not need to match and can contain completely different code. When opening a document, Ofice will use the compiled code if the code was compiled for the same version of Office on the same architecture. This means you can have a document with compiled code that does something bad while the source code looks completely safe. This technique is known as "VBA stomping".

This issue is made worse by the fact that worse that digital signatures on office documents do not apply to the compiled code as documented by [Dider Stevens](https://blog.nviso.eu/2020/06/04/tampering-with-digitally-signed-vba-projects/). An Ofice document with a signed macro can have it's compiled code tampered with and the digital signature will not be invalided.

## Potential Solution
Macro Exorcist can change the version value of a VBA macro in a Word document to the tentionally invalid value of 0xBEEF. When Word opens an "exorcised" document it will think the compiled code is for another version of Office and recompile the macro before executing it. The "exorcised" document can be freely edited and saved and the invalid version value will be retained unless the macro itself it editied (i.e. derivative documents of an "exorcised" document will remain "exorcised").

## Caveats
Modifying the version of VBA macro will invalidate the digital signature of the document (i.e. will prevent the macro from executing completely). This means a document needs to be "exorcised" prior to be being digitally signed (i.e. this cannot be used to make historical documents safe unless you also resign them). However, once a document has been "exorcised" and signed then all derivative documents based on it will also be "exorcised" and safe to open.

## Example Usage
Invalidating VBA project version:
```donthighlight
C:\Users\User\Desktop>MacroExorcist.exe evil.docm
[+] Saving backup to: evil.docm.bak
[+] Opening document...
[.] Word document recognised
[+] Processing document...
[+] Extracting VBA Project...
[.] VBA project version is : B200 - Office 2016/2019 (x64)
[+] Using BEEF to exorcise VBA demons... \m/
[+] Saving document...
[+] Exorcism successful!

```

Checking an already "exorcised" document:
```donthighlight
C:\Users\User\Desktop>MacroExorcist.exe evil.docm
[+] Saving backup to: evil.docm.bak
[+] Opening document...
[.] Word document recognised
[+] Processing document...
[+] Extracting VBA Project...
[.] VBA project version is: BEEF
[!] Document is already exorcised

```