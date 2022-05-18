; ============================================================
; Example
;   Links:
;       https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shfileoperationa
;       https://docs.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-shfileopstructa
;
;   Usage:
;       - [[ obj := FileOps() ]] ... to make an instance ...
;       - obj.Copy(src, dest [, no_confirm := false])
;       - obj.Rename & obj.Move same as above
;       - obj.Delete(src [, no_confirm := false])
;
;   Features:
;       - Adjustments for Unicode/ANSI and x86/x64
;       - "no_confirm := true" acts as "overwrite" on Copy action.
;
;       - Flags can be set directly by one of the following:
;           - obj.Flags := 0x1234
;           - obj.FlagStr := "flag_name flag_name ..."
;
;       - obj.error contains error codes on return.  Non-zero = an error.
;       - obj.abort contains abort codes in case of user/system aborting the action.
;
;   Limitations:
;       - Long paths are only supported in theory.  My tests indicate it doesn't work.
;       - Setting ProgressTitle doesn't seem to work.  The default title is more useful anyway.
;       - I don't plan to implement any class solutions for WantMappingHandle for now.
;       - Wildcards are not allowed in the directory name, an error will be thrown.
;       - All relative paths are assumed to be relative to A_ScriptDir and are converted to
;         fully qualified paths before execution.
;       - All flags are reset after execution.  This is done to prevent unexpected/dangerous results related
;         to mismanagement of flag values.
;       - If you set flags manually, then the [no_confirm] parameter is ignored.
;       
; ============================================================

; #INCLUDE D:\UserData\AHK\_INCLUDE\_LibraryV2\_JXON.ahk

; (FileOps) ; Init class for example.  I suggest using #INCLUDE near the top of your script instead.

; fso := FileOps()
; fso.FlagStr := "WANTMAPPINGHANDLE"
; q := Chr(34)

; Msgbox "Making a temp dir with "              ; Making Temp files
     ; . "temp files for testing..."            ; =================
; make_tmp_files()

; Msgbox "Testing Copy action..."
; fso.Copy("test_dir", "test_dir2")             ; Copy Example
; fso.Copy("test_dir", "test_dir3")             ; --- Setting up move example
; Msgbox "Showcasing confirmation dialog on collision (for copy action)..."
; fso.Copy("test_dir\test_1*", "test_dir2")     ; =================

; Msgbox "Now testing Delete action..."         ; Delete Example
; fso.Flags := 0                                ; --- Without this, files go to the Recycle Bin.
; fso.Delete("test_dir")                        ; =================

; Msgbox "Now testing Move action..."           ; Move Example      ; Acts as "Rename" when dest does not exist.
; fso.Move("test_dir2", "test_dir3")            ; ================= ; Otherwise it moves src dir into dest dir.
; MsgBox "Now test_dir2 is inside test_dir3 folder.`n`n"

; MsgBox "Now a Rename example."                ; Rename Example
; fso.Rename("test_dir3", "test_dir4")          ; =================

; Msgbox "Now cleaning up.`r`n`r`n"             ; Setting manual flag values.
     ; . "Nukeing test_dir4."                   ; By default, all deleted files will go to the recycle bin except in
; fso.Flags := 0 ; no AllowUndo, nuke the files ; unusual cases.  If you want notification of these cases then add
; fso.Delete("test_dir4")                       ; the WantNukeWarning flag >> fso.Flags := WantNukeWarning | other_values

; make_tmp_files() {                            ; You can also set flag values like this:
    ; DirCreate("test_dir")                     ; fso.FlagStr := "WantNukeWarning flag_name flag_name ..."
    ; Loop 50                                   ; Take care in the flags that you use AND don't use.
        ; FileAppend "test" A_Index, "test_dir\test_" A_Index ".txt"
; }

; ============================================================
; FileOps class
;       - obj.Copy(src, dest [, no_confirm := false])
;       - obj.Rename & obj.Move same as above
;       - obj.Delete(src [, no_confirm := false])
;
;       - "no_confirm := true" acts as "overwrite" on Copy action.
;       - 
;       - Flags can be set directly by one of the following:
;           - obj.Flags := 0x1234
;           - obj.FlagStr := "flag_name flag_name ..."
;           ** WARNING ** DO NOT SET BOTH PROPERTIES - THE BEHAVIOR OF SETTING BOTH IS UNDEFINED
;
;       - obj.error contains error codes on return.  Non-zero = an error.
;       - obj.abort contains abort codes in case of user/system aborting the action.
;
;       - obj.NameMappings contains a Map() of any files renamed on collison.
;         This requires that the following flags be set:
;               RENAMEONCOLLISION
;               WANTMAPPINGHANDLE
;
;         You can check for name mappings like this:
;
;               If obj.NameMappings.Count
;                   ; ... do stuff
;
;         You can iterate through the list like this:
;
;               For oldPath, newPath in obj.NameMappings {
;                   ; ... do stuff
;               }
;
; ============================================================

class FileOps {
    Static hwnd := 0     ; SHFILEOPSTRUCT
         , wFunc         := (A_PtrSize=4) ? 4  : 8
         , From          := (A_PtrSize=4) ? 8  : 16
         , To            := (A_PtrSize=4) ? 12 : 24
         , Flags         := (A_PtrSize=4) ? 16 : 32
         , AnyOpsAbort   := (A_PtrSize=4) ? 18 : 36
         , NameMappings  := (A_PtrSize=4) ? 22 : 40
         , ProgressTitle := (A_PtrSize=4) ? 26 : 48
    
    Static u := StrLen(Chr(0xFFFF)) ; IsUnicode
         , _func := (this.u) ? "SHFileOperationW" : "SHFileOperationA"
         , fo := {Move:1, Copy:2, Delete:3, Rename:4}
    
    Static oldPath := 0 ; SHNAMEMAPPING
         , newPath := (A_PtrSize=4) ? 4 : 8
         , opcc    := (A_PtrSize=4) ? 8 : 16
         , npcc    := (A_PtrSize=4) ? 12 : 20
    
    ; Static LongPaths := RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\FileSystem","LongPathsEnabled") ; maybe...
    
    Static FILEOP_FLAGS := {MultiDestFiles:0x1     
                           ,Silent:0x4                      ; Do not display progress dialog.
                           ,RenameOnCollision:0x8           ; Rename instead of overwrite or prompt.
                           ,NoConfirmation:0x10             ; Respond with Yes to All for any dialog box that is displayed.
                           ,WantMappingHandle:0x20          ; Track file renaming ... i might not implement this.
                           ,AllowUndo:0x40                  ; Preserve undo information, if possible.  AKA Recycle Bin.
                           ,FilesOnly:0x80                  ; Only on files if a wildcard file name (.) is specified.
                           ,SimpleProgress:0x100            ; Unexpanded progress GUI.
                           ,NoConfirmMkDir:0x200            ; Dont' ask user to confirm on MKDIR if the operation requires it.
                           ,NoErrorUI:0x400                 ; Hide errors from user.
                           ,NoUI:0x614                      ; Same as = (Silent | NoConfirmation | NoErrorUI | NoConfirmMkDir)
                           ,NoCopySecurityAttribs:0x800     ; Do not copy the security attributes of the file.
                           ,NoRecursion:0x1000              ; Current dir only.
                           ,NoConnectedElements:0x2000      ; Do not move connected files as a group.
                           
                           ,WantNukeWarning:0x4000}         ; Send a warning if a file is being permanently destroyed during
                                                            ; a delete operation rather than recycled. This flag partially 
                                                            ; overrides FOF_NOCONFIRMATION.
    
    FlagStr := "", Flags := "", error := 0, abort := 0, Title := ""
    
    __Call(name, p) {
        msg1 := "Minimum required parameters is " ((name!="Delete")?"two":"one") ":`r`n`r`n"
              . "Usage:`r`n`r`n"
              . "obj." name "(From, " ((name!="Delete")?"To":"") " [, no_confirm := false])"
        msg2 := "The first two parameters must be strings:`r`n`r`n"
              . "Usage:`r`n`r`n"
              . "obj." name "(From, " ((name!="Delete")?"To":"") " [, no_confirm := false])"
        
        If (name != "Delete")
            If (p.Length < 2) || (p.Length > 3)
                throw Error("Invalid parameters.",, msg1)
            Else If  (Type(p[1]) != "String") || (Type(p[2]) != "String")
                throw Error("Invalid parameter type.",, msg2)
        Else
            If (p.Length < 1) || (p.Length > 2)
                throw Error("Invalid parameters.",, msg1)
            Else If  (Type(p[1]) != "String")
                throw Error("Invalid parameter type.",, msg2)
        
        this.nameMappings := Map()
        this.Operation(name, p)
    }
    
    Operation(name, p) {
        Static flags := FileOps.FILEOP_FLAGS
        
        If (name != "Delete")
            str_from   := p[1]
          , str_to     := p[2]
          , no_confirm := (p.Has(3)?p[3]:false)
        Else
            str_from   := p[1]
          , no_confirm := (p.Has(2)?p[2]:false)
        
        obj := FileOps.SHFILEOPSTRUCT()
        obj.wFunc := FileOps.fo.%name%  ; function = Copy / Move / Delete / Rename
        str_from  := this.ValidatePath(str_from)                        ; validate paths
        (name!="Delete") ? (str_to := this.ValidatePath(str_to)) : ""   ; validate paths
        
        If (this.Flags="" && this.FlagStr="")
            this.Flags := flags.AllowUndo | flags.SimpleProgress
          , (no_confirm) ? this.Flags := this.Flags | flags.NoConfirmation : ""  ; NoConfirm = Overwrite (mostly)
        
        If (this.FlagStr && this.Flags="") {
            this.Flags := 0
            arr := StrSplit(this.FlagStr," ")
            For i, _flag in arr
                If flags.HasProp(_flag)
                    this.Flags := this.Flags | flags.%_flag%
        }
        
        (this.Title) ? (obj.ProgressTitle := this.Title) : ""
        
        If (name!="Delete") && InStr(str_to,"`n") && !(this.Flags & flags.MultiDestFiles)   ; If obj.To is multi-line,
            this.Flags := this.Flags | flags.MultiDestFiles                                 ; this should be automatic...
        
        obj.From := (bFrom := this.MkStr(str_from)).ptr
        , obj.Flags := this.Flags
        If (name!="Delete")
            bTo := this.MkStr(str_to)
          , obj.To := bTo.ptr
        
        this.error := DllCall("shell32\" FileOps._func, "UPtr", obj.ptr)    ; store error code
        this.abort := obj.AnyOpsAbort                                       ; store abort code
        this.GetNameMappings(obj.NameMappings)
        ; msgbox jxon_dump(this.NameMappings,4)
        
        this.Flags := "", this.FlagStr := "", this.Title := "" ; reset flags after operation
        return !this.error
    }
    ValidatePath(_in) { ; convert relative to absolute, and check for dangerous paths/errors
        Static flags := FileOps.FILEOP_FLAGS
        result := "", multi := false
        Loop Parse _in, "`n", "`r"
        {
            If (path := Trim(A_LoopField," `t`\")) {    ; trim trailing or beginning "\"
                path := StrReplace(path,"/","\")        ; fix "/" to "\"
                If !RegExMatch(path, "i)^[a-z]\:")      ; if not an absolute path, append A_ScriptDir
                    path := A_ScriptDir "\" path
                
                SplitPath path,, &test_dir
                If RegExMatch(test_dir,"[\*\?]")
                    throw Error("Wildcard (*) characters are not allowed in the directory name.  "
                              . "Only use wildcard characers at the end of the path, in the file part.")
                
                result .= (result?"`r`n":"") path ; long paths -- ; (FileOps.u?"\\?\":"") -- (unicode only)... tempermental
            }
        }
        
        return result
    }
    MkStr(_in, dNull:=true) { ; returns string buffer
        Static mStr := (FileOps.u)?2:1 ; string multiplier (unicode vs ansi)
        
        bSize := 0
        Loop Parse (_in := Trim(_in,"`r`n")), "`n", "`r"
            If (A_LoopField)
                bSize += ((StrLen(A_LoopField) * mStr) + mStr) ; str len + NULL ender
        bSize += (dNull)?mStr:0 ; add double NULL at very end, by default
        
        buf := Buffer(bSize, 0), offset := 0
        Loop Parse _in, "`n", "`r"
        {
            StrPut(A_LoopField, buf.ptr + offset)
            offset += ((StrLen(A_LoopField) * mStr) + mStr)
        }
        
        return buf
    }
    GetNameMappings(ptr) {
        this.NameMappings := Map()
        If ptr && (count := NumGet(ptr,"UPtr")) {
            nm := FileOps.SHNAMEMAPPING(NumGet(ptr,A_PtrSize,"UPtr"))
            Loop count
                this.NameMappings[nm.oldPath] := nm.newPath, nm.ptr += nm.size
            r := DllCall("shell32\SHFreeNameMappings","UPtr",ptr) ; 
        }
    }
    
    class SHFILEOPSTRUCT {
        __New() {
            fo := FileOps
            this.struct := Buffer((A_PtrSize=4)?30:56, 0)
            this.MkStr := FileOps.Prototype.MkStr
            
            this.DefineProp("ptr",{Get:(*)=>this.struct.ptr})
            list := Map("hwnd","UPtr","wFunc","UInt","From","UPtr","To","UPtr","Flags","UShort","AnyOpsAbort","Int","NameMappings","UPtr")
            For _name, _type in list
                this.DefineProp(_name,{Get:this._get.Bind(,FileOps.%_name%,_type), Set:this._set.Bind(,FileOps.%_name%,_type)})
        }
        _get(offset,_type) => NumGet(this.struct, offset, _type)
        _set(offset,_type,value) => NumPut(_type, value, this.struct, offset)
        
        ProgressTitle { ; this doesn't work as hoped.
            get => (ptr := NumGet(this.struct, FileOps.ProgressTitle, "UPtr")) ? StrGet(ptr) : ""
            set {
                this.titleBuf := this.MkStr(value, false)
                NumPut("UPtr", this.titleBuf.ptr, this.struct, FileOps.ProgressTitle)
            }
        }
    }
    
    class SHNAMEMAPPING {
        __New(ptr:=0) {
            this.size := (A_PtrSize=4) ? 16 : 24
            this.ptr := ptr
            this.DefineProp("opcc",{Get:(*)=>NumGet(this.ptr, FileOps.opcc, "Int")})
            this.DefineProp("npcc",{Get:(*)=>NumGet(this.ptr, FileOps.npcc, "Int")})
        }
        oldPath => StrGet(NumGet(this.ptr,FileOps.oldPath,"UPtr"), this.opcc)
        newPath => StrGet(NumGet(this.ptr,FileOps.newPath,"UPtr"), this.npcc)
    }
}



; typedef struct _SHFILEOPSTRUCTA {          offset     size [32/64]
  ; HWND         hwnd;                      | 0          4/8
  ; UINT         wFunc;                     | 4/8        8/12
  ; PCZZSTR      pFrom;                     | 8/16 <--- 12/24 --- x64 offset
  ; PCZZSTR      pTo;                       |12/24      16/32
  ; FILEOP_FLAGS fFlags;                    |16/32      18/34
  ; BOOL         fAnyOperationsAborted;     |20/36 <--- 24/40 --- offset for BOOL
  ; LPVOID       hNameMappings;             |24/40      28/48
  ; PCSTR        lpszProgressTitle;         |28/48      32/56
; } SHFILEOPSTRUCTA, *LPSHFILEOPSTRUCTA;

; =======================================================================================
; tests with GetFullPathName
; =======================================================================================
; path := "test3.txt"
; sz := DllCall("GetFullPathName", "Str", path, "UInt", 0, "UPtr", 0, "UPtr", 0)

; msgbox "size: " sz " / " StrPut(path)

; buf := Buffer(sz*2, 0)
; sz := DllCall("GetFullPathName", "Str", path, "UInt", sz*2, "UPtr", buf.ptr, "UPtr", 0)

; msgbox "size: " sz "`r`n"
     ; . "test: " StrGet(buf)

; dbg(_in) { ; AHK v2
    ; Loop Parse _in, "`n", "`r"
        ; OutputDebug "AHK: " A_LoopField
; }