use nu_plugin::{serve_plugin, MsgPackSerializer};

use nu_plugin_to_xlsx::XLSX;

fn main() {
    serve_plugin(&mut XLSX {}, MsgPackSerializer {});
}