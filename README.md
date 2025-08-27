# nu_plugin_to_xlsx

A Nushell plugin to export Nushell data to XLSX files.

## Installation

### Build from source

```bash
cargo build --release
plugin add target/release/nu_plugin_to_xlsx[.exe]
plugin use to_xlsx
```

### Install with cargo

```bash
cargo install nu_plugin_to_xlsx
```

### Register the plugin

To add the plugin from inside Nushell:

```bash
plugin add ~/.cargo/bin/nu_plugin_to_xlsx
plugin use to_xlsx
```

## Usage

### Save a record to XLSX file with sheet name

```nushell
{user: bob, age: 30} | to xlsx User.xlsx
```

### Save a table to Excel XLSX file

```nushell
echo [[name]; [bob]] | to xlsx
```

### Save piped data to XLSX file with sheet name

```nushell
ls | to xlsx Files.xlsx
```

### Save list of records to XLSX file with sheet name

```nushell
[{user: bob, age: 30}, {user: john, age: 40}] | to xlsx Users.xlsx
```

### save a nested table to XLSX file

```nushell
[
  {
    "age": 24,
    "files": [
      {
        "name": "Cargo.lock",
        "type": "file",
        "size": 51715,
        "modified": "2024-12-09 07:27:47.260591326 +08:00"
      },
      {
        "name": "Cargo.toml",
        "type": "file",
        "size": 260,
        "modified": "2024-12-09 07:27:39.112955909 +08:00"
      },
      {
        "name": "a.xlsx",
        "type": "file",
        "size": 5830,
        "modified": "2024-12-14 17:08:21.029243394 +08:00"
      },
      {
        "name": "b.xlsx",
        "type": "file",
        "size": 6277,
        "modified": "2024-12-09 19:24:08.426020860 +08:00"
      },
      {
        "name": "src",
        "type": "dir",
        "size": 96,
        "modified": "2024-12-09 18:55:30.097278053 +08:00"
      },
      {
        "name": "target",
        "type": "dir",
        "size": 192,
        "modified": "2024-12-09 06:54:15.325554495 +08:00"
      }
    ],
    "name": "john",
    "signed_up": false
  },
  {
    "age": 30,
    "files": [
      {
        "name": "Cargo.lock",
        "type": "file",
        "size": 51715,
        "modified": "2024-12-09 07:27:47.260591326 +08:00"
      },
      {
        "name": "Cargo.toml",
        "type": "file",
        "size": 260,
        "modified": "2024-12-09 07:27:39.112955909 +08:00"
      },
      {
        "name": "src",
        "type": "dir",
        "size": 96,
        "modified": "2024-12-09 18:55:30.097278053 +08:00"
      },
      {
        "name": "target",
        "type": "dir",
        "size": 192,
        "modified": "2024-12-09 06:54:15.325554495 +08:00"
      }
    ],
    "name": "mike",
    "signed_up": true
  },
{
    "age": 30,
    "files": [
      {
        "name": "Cargo.lock",
        "type": "file",
        "size": 51715,
        "modified": "2024-12-09 07:27:47.260591326 +08:00"
      },
      {
        "name": "Cargo.toml",
        "type": "file",
        "size": 260,
        "modified": "2024-12-09 07:27:39.112955909 +08:00"
      },
      {
        "name": "src",
        "type": "dir",
        "size": 96,
        "modified": "2024-12-09 18:55:30.097278053 +08:00"
      },
      {
        "name": "target",
        "type": "dir",
        "size": 192,
        "modified": "2024-12-09 06:54:15.325554495 +08:00"
      }
     
    ],
    "name": "mike",
    "signed_up": true
  }
] | to xlsx NestedFiles.xlsx
```