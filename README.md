<div align="center">

## Rename, copy, delete entire registry keys


</div>

### Description

This is a registry API module that contains the usual functions to create, set, delete and count registry keys and values. Since the API does not provide a way to rename registry keys, the only way to do it is to copy the entire key including all subkeys and values which could be of different types, then delete the old key recursively (which is also not supported by the API). These functions are also included in this module, the rename/copy function supports all existing value types (REG_SZ, REG_EXPAND_SZ, REG_BINARY, REG_DWORD, REG_MULTI_SZ). The delete function has a security feature (minimum depth) to avoid accidental recursive deletion of important keys, such as HKLM\Software, etc. Make sure you read the comments I left to understand how this function works. The only other approach to renaming a registry key (which also works by copying, etc.) involves using a function to write all keys to a temp file and then save them under the new name. There is code available on the internet that demostrate this approach, but you will always need a temp file, which is not good. This makes my code, at least as of today and as far as I know, one of a kind. I hope you find it useful.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2006-09-08 23:41:58
**By**             |[Luis Cantero](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/luis-cantero.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Registry](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/registry__1-36.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Rename\_\_co2019429122006\.zip](https://github.com/Planet-Source-Code/luis-cantero-rename-copy-delete-entire-registry-keys__1-66540/archive/master.zip)








