<div align="center">

## String, array and file compression with Zlib\.dll


</div>

### Description

This code has been upgraded in another post. Please look for 'Business-strength compression, encryption and signing (single and public/private keys via CryptoAPI)' on this site.

This wrapper class and example program demonstrates the proper use of the Zlib compression dll. Unbelievably, Windows does not provide adequate string/file compression in the API (only decompression of files created with the compress.exe or compact.exe utilities that come with Windows, and their compression is weak by today's standards). No string or array compression support...and having to call an exe to compress a file is far from perfect.

To fill this void a group wrote the Zlib.dll utility (included; their web site address is included, too, so you can get the latest version). They used the same code to produce the compression algorithms that are used by the newer PNG picture format. So if you have used PNG you have used Zlib. I have seen the Zlib dll used by several commercial programs, too.

It provides fast and compact encryption for byte arrays. This wrapper class extends it to strings and full files, too. Two other postings on planetsourcecode deal with Zlib. One is barely functional and the other only provides the coverted C header files (but is otherwise excellent with many utilities and modules to do other things -- check out "Kira" posted by The_Lung). I found this code on the Zlib web site, and converted it from an ocx to a regular class module to conserve resources. I also moved all the code dealing with file compression into the class where it belonged.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-11-28 11:28:16
**By**             |[Doug Gaede](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/doug-gaede.md)
**Level**          |Intermediate
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD1216411282000\.zip](https://github.com/Planet-Source-Code/doug-gaede-string-array-and-file-compression-with-zlib-dll__1-13184/archive/master.zip)








