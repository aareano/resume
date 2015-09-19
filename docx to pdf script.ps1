$Word = NEW-OBJECT -COMOBJECT WORD.APPLICATION
$file = Get-ChildItem "Resume.docx"
$Doc = $Word.Documents.Open($file.FullName)
$Doc.saveas([ref] (($Doc).FullName.Replace("docx","pdf")), [ref] 17)
$Doc.close()