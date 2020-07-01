# PDFtoXLS
Convert PDF email attachment to XLS files in VBA

Not as simple as it seems.

The Error checking is at the minimal. It could do with more advance checking.

It does not overwrite files, in its current state it will ask if you want to overwrite

Big issue:

  If the Macro crashes for any reason, you might be left with Word and/or Excel running in the background.
  Therefore you might fid you cannot delete the 'pdf' and/or 'xlsx'
  
