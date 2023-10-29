# Xlsx plugin for Nushell
A nushell plugin to export nushel data to xlsx file 

To install:

```
> cargo install --path .
```

To register (from inside Nushell):
```
> register <path to installed plugin> 
```

Usage:
```
<table_data> | to xlsx  <file_name> [sheet_name]
```


