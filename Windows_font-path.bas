' c:\lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -excel -wsh "Windows_font-path" -c FontPath
'
'   --> https://github.com/ReneNyffenegger/runVBAFilesInOffice/blob/master/runVBAFilesInOffice.vbs
'
public sub FontPath()

  dim wsh as new WshShell
  
  cells(1,1) = "Font Path:"
  cells(1,2) = wsh.specialFolders("Fonts")

  activeWorkbook.saved = true

end sub
