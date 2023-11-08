# Xlsx plugin for Nushell
A nushell plugin to export nushel data to xlsx file 


## build from source (for now )

```
cargo build --release
register target/release/nu_plugin_xlsx[.exe]
```

## Install with cargo (later):

- Install 
```
cargo install nu_plugin_xlsx
```

- To register (from inside Nushell):

```
> register ~/.cargo/bin/nu_plugin_xlsx
```

##  Usage 

  Save a record to xlsx file with sheet name
  > {user: bob, age: 30} | to xlsx User.xlsx

  Save a table to excel xlsx file
  > echo [[name]; [bob]] | to xlsx

  Save piped data to xlsx file with sheet name
  > ls | to xlsx Files.xlsx

  Save list of record to to xlsx file with sheet name
  > [{user: bob, age: 30},{user: john, age:40}] | to xlsx Users.xlsx
