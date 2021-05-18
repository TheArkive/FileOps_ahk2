# FileOps_ahk2
File Copy, Delete, Rename, Move with default Win32 progress dialog.

### Links:
* [SHFileOperationA Function](https://docs.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shfileoperationa)
* [SHFILEOPSTRUCTA Structure](https://docs.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-shfileopstructa)

### Usage:
* `obj := FileOps()` to make an instance
* `obj.Copy(src, dest [, no_confirm := false])`
* `obj.Rename()` & `obj.Move()` same as above
* `obj.Delete(src [, no_confirm := false])`

### Features:
* Adjustments for Unicode/ANSI and x86/x64
* `no_confirm := true` acts as "overwrite" on Copy action.
* Flags can be set directly by one of the following:
  * `obj.Flags := 0x1234`
  * `obj.FlagStr := "flag_name flag_name ..."`
* `obj.error` contains error codes on return.  Non-zero = an error.
* `obj.abort` contains abort codes in case of user/system aborting the action.

### Limitations:
* Long paths are only supported in theory.  My tests indicate it doesn't work.
* Setting ProgressTitle doesn't seem to work.  The default title is more useful anyway.
* I don't plan to implement any class solutions for WantMappingHandle for now.
* Wildcards are not allowed in the directory name, an error will be thrown.
* All relative paths are assumed to be relative to `A_ScriptDir` and are converted to fully qualified paths before execution.
* All flags are reset after execution.  This is done to prevent unexpected/dangerous results related to mismanagement of flag values.
* If you set flags manually, then the `no_confirm` parameter is ignored.