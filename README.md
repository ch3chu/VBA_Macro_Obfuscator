# VBA_Macro_Obfuscator
Tool to create obfuscated VBA code for macros in Office documents.
This script creates two vba files (final.vba and fake.vba) to use the “VBA Stomping” technique.

# Usage
At the moment this script only supports .exe files.

```
> python obfus_vba.py -h
usage: obfus_vba.py [-h] -f FILE -t {exe}

VBA Obfuscator

options:
  -h, --help            show this help message and exit
  -f FILE, --file FILE  File to be embedded
  -t {exe}, --type {exe}
                        File type
```

1. Generates the two VBA files: `python3 obfus_vba.py -f calc.exe -t exe`
2. Create an office document, click View -> Macros -> Create
3. Paste de fake.vba code into the ThisDocument macro and save the document
4. You can use EvilClippy to stomp the vba code like this: `EvilClippy.exe -s final.vba Doc1.doc`
5. You can use OfficePurge to purge P-code from module streams: `OfficePurge.exe -d word -f Doc1_EvilClippy.doc -m ThisDocument`
