' Silinecek komut adı
komutAdi = "FrontEnd-Setup"

' Windows Kayıt Defteri'nde değişiklik yapmak için kullanılacak komut
regSilKomutu = "reg delete ""HKEY_CLASSES_ROOT\Directory\Background\shell\" & komutAdi & """ /f"

' Windows Komut İstemcisini yönetici olarak başlat
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cmd.exe", "/c " & regSilKomutu, "", "runas", 1