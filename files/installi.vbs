' Windows Kayıt Defteri'ne yeni bir komut ekleyen VBScript

' Komut adı ve komut dosyasının yolu
komutAdi = "FrontEnd-Setup"
komutDosyasi = "C:\FrontEnd-Setup.bat"

' Windows Kayıt Defteri'nde değişiklik yapmak için kullanılacak komut
regKomutu = "reg add ""HKEY_CLASSES_ROOT\Directory\Background\shell\" & komutAdi & """ /ve /t REG_SZ /d """ & komutAdi & """ /f"
regKomutu2 = "reg add ""HKEY_CLASSES_ROOT\Directory\Background\shell\" & komutAdi & "\command"" /ve /t REG_SZ /d """ & komutDosyasi & """ /f"

' Windows Komut İstemcisini yönetici olarak başlat
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cmd.exe", "/c " & regKomutu & " & " & regKomutu2, "", "runas", 1