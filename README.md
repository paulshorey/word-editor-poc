## Goal:

This attempts to insert the contents of another DOCX file into the current DOCX file,

by creating a "content control" in the current DOCX file, then use **`cc.insertFileFrombase64`**

to insert a `base64`.

## Unfortunately:

It does not work well. The file does get inserted every time (from base64 string),

but calling the `.insertFileFromBase64` breaks the MS Word editor app UI. The editor

shows a "Waiting..." then "Inserting..." popup, but it never goes away. It gets stuck!

## Debugging:

Tried everything. We can use OOXML too instead of base64, but getting the same error. :(

We even tried using the API to output the contents of a "content control" as OOXML,

then use that exact correct OOXML with **`cc.insertOoxml`** to insert the file. But no bueno.

## Resources:

Convert word files to base64 string:
https://products.aspose.app/pdf/conversion/docx-to-base64
