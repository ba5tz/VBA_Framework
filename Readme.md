# VBA Framework

Dibuat dengan tujuan untuk pembuatan aplikasi berbasis VBA yang lebih mudah dan ramah bagi pemula

```VBA
Dim This as New MyFrameWork
```

Set Database dan kolom
```VBA
 This.SetDatabase Sheet1, "A"    
```

Set Control 
```VBA
 This.setControl = Textbox1
 This.setControl = Combobox1
```

Reset Control (Membersihkan Control) 
```VBA
 This.Reset
```

Baris terkahir dari Set Database
```VBA
Dim Baris as long 
baris = This.gerbarisDB
```

Baris terkahir dari Sheet dan Kolom Bebas
```VBA
Dim Baris as long 
baris = This.gerbaris(Sheet1, "A")
```

