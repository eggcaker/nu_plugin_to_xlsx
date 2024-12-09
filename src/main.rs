use nu_plugin::{serve_plugin, Plugin, PluginCommand, EvaluatedCall, EngineInterface};
use nu_protocol::{Category, LabeledError, PipelineData, Signature, SyntaxShape, Type, Value};
use rust_xlsxwriter::{Workbook, Worksheet, Format, Color};
use std::path::PathBuf;

struct ToXlsx;

#[derive(Default)]
struct NestedTableInfo {
    size: usize,
    column_count: usize,
}

impl ToXlsx {
    fn new() -> Self {
        Self
    }

    fn write_value(&self, workbook: &mut Workbook, worksheet_name: &str, value: &Value) -> Result<(), Box<dyn std::error::Error>> {
        let mut worksheet = workbook.add_worksheet().set_name(worksheet_name)?;
        let header_format = Format::new()
            .set_bold()
            .set_background_color(Color::RGB(0x202020))
            .set_font_color(Color::RGB(0x00FF00));

        match value {
            Value::Record { val, .. } => {
                // Write headers
                worksheet.write_string_with_format(0, 0, "Key", &header_format)?;
                worksheet.write_string_with_format(0, 1, "Value", &header_format)?;

                // First pass: collect nested table info
                let mut nested_tables = Vec::new();
                for (_, value) in val.iter() {
                    if let Value::List { vals, .. } = value {
                        if let Some(Value::Record { val: first_record, .. }) = vals.first() {
                            nested_tables.push(NestedTableInfo {
                                size: vals.len(),
                                column_count: first_record.len(),
                            });
                        } else {
                            nested_tables.push(NestedTableInfo::default());
                        }
                    } else {
                        nested_tables.push(NestedTableInfo::default());
                    }
                }

                // Calculate column offsets based on nested table widths
                let mut column_offsets = Vec::new();
                let mut current_offset = 0;
                for (_, value) in val.iter() {
                    if let Value::List { vals, .. } = value {
                        if let Some(Value::Record { val: first_record, .. }) = vals.first() {
                            column_offsets.push(current_offset);
                            current_offset += first_record.len();
                        } else {
                            column_offsets.push(current_offset);
                            current_offset += 1;
                        }
                    } else {
                        column_offsets.push(current_offset);
                        current_offset += 1;
                    }
                }

                // Calculate row offsets for vertical spacing
                let mut row_offsets = Vec::with_capacity(nested_tables.len());
                let mut offset = 0;
                for info in &nested_tables {
                    row_offsets.push(offset);
                    offset += info.size + 1;  // Add 1 for the header row
                }

                // Second pass: write data with proper spacing
                let mut current_row = 1;
                for ((((key, value), _table_info), row_offset), col_offset) in val.iter()
                    .zip(nested_tables.iter())
                    .zip(row_offsets.iter())
                    .zip(column_offsets.iter())
                {
                    let actual_row = current_row + (*row_offset as u32);
                    worksheet.write_string(actual_row, *col_offset as u16, key)?;

                    match value {
                        Value::List { vals, .. } => {
                            if let Some(Value::Record { val: first_record, .. }) = vals.first() {
                                // Get headers from first record
                                let headers: Vec<String> = first_record.columns().into_iter().map(|s| s.to_string()).collect();

                                // Write nested table headers
                                for (col, header) in headers.iter().enumerate() {
                                    worksheet.write_string_with_format(actual_row, (*col_offset + col) as u16, header, &header_format)?;
                                }

                                // Write nested table data
                                for (nested_row, record_value) in vals.iter().enumerate() {
                                    if let Value::Record { val, .. } = record_value {
                                        for (col, header) in headers.iter().enumerate() {
                                            if let Some(cell_value) = val.get(header) {
                                                self.write_cell_value(&mut worksheet, actual_row + 1 + (nested_row as u32), (*col_offset + col) as u16, cell_value)?;
                                            }
                                        }
                                    }
                                }

                                current_row += 1;  // Move to next row after this nested table
                            } else {
                                worksheet.write_string(actual_row, (*col_offset + 1) as u16, &format!("{:?}", vals))?;
                                current_row += 1;
                            }
                        }
                        _ => {
                            self.write_cell_value(&mut worksheet, actual_row, (*col_offset + 1) as u16, value)?;
                            current_row += 1;
                        }
                    }
                }
            }
            Value::List { vals, .. } => {
                if let Some(Value::Record { val: first_record, .. }) = vals.first() {
                    // Write headers
                    let headers: Vec<String> = first_record.columns().into_iter().map(|s| s.to_string()).collect();
                    for (col, header) in headers.iter().enumerate() {
                        worksheet.write_string_with_format(0, col as u16, header, &header_format)?;
                    }

                    // Calculate column offsets based on nested table widths
                    let mut column_offsets = Vec::new();
                    let mut current_offset = 0;
                    for record_value in vals.iter() {
                        if let Value::Record { val, .. } = record_value {
                            for (_, value) in val.iter() {
                                if let Value::List { vals, .. } = value {
                                    if let Some(Value::Record { val: nested_record, .. }) = vals.first() {
                                        column_offsets.push(current_offset);
                                        current_offset += nested_record.len();
                                    } else {
                                        column_offsets.push(current_offset);
                                        current_offset += 1;
                                    }
                                } else {
                                    column_offsets.push(current_offset);
                                    current_offset += 1;
                                }
                            }
                        }
                    }

                    // First pass: collect nested table info
                    let mut nested_tables = Vec::new();
                    for record_value in vals.iter() {
                        if let Value::Record { val, .. } = record_value {
                            let mut row_info = NestedTableInfo::default();
                            for (_, value) in val.iter() {
                                if let Value::List { vals, .. } = value {
                                    if let Some(Value::Record { val: nested_record, .. }) = vals.first() {
                                        row_info = NestedTableInfo {
                                            size: vals.len(),
                                            column_count: nested_record.len(),
                                        };
                                        break;
                                    }
                                }
                            }
                            nested_tables.push(row_info);
                        }
                    }

                    // Calculate row offsets for vertical spacing
                    let mut row_offsets = Vec::with_capacity(nested_tables.len());
                    let mut offset = 0;
                    for info in &nested_tables {
                        row_offsets.push(offset);
                        offset += info.size + 1;  // Add 1 for the header row
                    }

                    // Second pass: write data with proper spacing
                    let mut current_row = 1;
                    for ((record_value, _table_info), row_offset) in vals.iter().zip(nested_tables.iter()).zip(row_offsets.iter()) {
                        let actual_row = current_row + (*row_offset as u32);
                        if let Value::Record { val, .. } = record_value {
                            for ((_col, (_header, cell_value)), col_offset) in headers.iter().zip(val.iter()).zip(column_offsets.iter()) {
                                match cell_value {
                                    Value::List { vals, .. } => {
                                        if let Some(Value::Record { val: first_record, .. }) = vals.first() {
                                            // Get nested headers
                                            let nested_headers: Vec<String> = first_record.columns().into_iter().map(|s| s.to_string()).collect();

                                            // Write nested headers
                                            for (nested_col, nested_header) in nested_headers.iter().enumerate() {
                                                worksheet.write_string_with_format(actual_row, (*col_offset + nested_col) as u16, nested_header, &header_format)?;
                                            }

                                            // Write nested data
                                            for (nested_row, nested_record) in vals.iter().enumerate() {
                                                if let Value::Record { val: nested_val, .. } = nested_record {
                                                    for (nested_col, nested_header) in nested_headers.iter().enumerate() {
                                                        if let Some(nested_cell_value) = nested_val.get(nested_header) {
                                                            self.write_cell_value(&mut worksheet, actual_row + 1 + (nested_row as u32), (*col_offset + nested_col) as u16, nested_cell_value)?;
                                                        }
                                                    }
                                                }
                                            }
                                        } else {
                                            self.write_cell_value(&mut worksheet, actual_row, *col_offset as u16, cell_value)?;
                                        }
                                    }
                                    _ => {
                                        self.write_cell_value(&mut worksheet, actual_row, *col_offset as u16, cell_value)?;
                                    }
                                }
                            }
                        }
                        current_row += 1;  // Move to next row after this nested table
                    }
                } else {
                    // Handle simple list
                    for (row, item) in vals.iter().enumerate() {
                        self.write_cell_value(&mut worksheet, row as u32, 0, item)?;
                    }
                }
            }
            _ => {
                // Write single value
                self.write_cell_value(&mut worksheet, 0, 0, value)?;
            }
        }

        Ok(())
    }

    fn write_cell_value(&self, worksheet: &mut Worksheet, row: u32, col: u16, value: &Value) -> Result<(), Box<dyn std::error::Error>> {
        match value {
            Value::String { val, .. } => {
                worksheet.write_string(row, col, val)?;
            }
            Value::Int { val, .. } => {
                worksheet.write_number(row, col, *val as f64)?;
            }
            Value::Float { val, .. } => {
                worksheet.write_number(row, col, *val)?;
            }
            Value::Bool { val, .. } => {
                worksheet.write_boolean(row, col, *val)?;
            }
            Value::Date { val, .. } => {
                worksheet.write_string(row, col, &format!("{}", val))?;
            }
            Value::Filesize { val, .. } => {
                let size_str = if *val < 1024 {
                    format!("{} B", val)
                } else if *val < 1024 * 1024 {
                    format!("{:.1} KB", *val as f64 / 1024.0)
                } else if *val < 1024 * 1024 * 1024 {
                    format!("{:.1} MB", *val as f64 / (1024.0 * 1024.0))
                } else {
                    format!("{:.1} GB", *val as f64 / (1024.0 * 1024.0 * 1024.0))
                };
                worksheet.write_string(row, col, &size_str)?;
            }
            Value::Duration { val, .. } => {
                let duration_str = if *val < 1_000 {
                    format!("{} ns", val)
                } else if *val < 1_000_000 {
                    format!("{:.1} µs", *val as f64 / 1_000.0)
                } else if *val < 1_000_000_000 {
                    format!("{:.1} ms", *val as f64 / 1_000_000.0)
                } else {
                    format!("{:.1} s", *val as f64 / 1_000_000_000.0)
                };
                worksheet.write_string(row, col, &duration_str)?;
            }
            Value::List { vals, .. } => {
                worksheet.write_string(row, col, &format!("{:?}", vals))?;
            }
            _ => {
                worksheet.write_string(row, col, &format!("{:?}", value))?;
            }
        }
        Ok(())
    }
}

#[derive(Clone)]
struct ToXlsxCommand;

impl PluginCommand for ToXlsxCommand {
    type Plugin = ToXlsx;

    fn name(&self) -> &str {
        "to xlsx"
    }

    fn signature(&self) -> Signature {
        Signature::build("to xlsx")
            .input_output_types(vec![(Type::Any, Type::Nothing)])
            .required("path", SyntaxShape::String, "Path to output xlsx file")
            .category(Category::Experimental)
    }

    fn description(&self) -> &str {
        "Export data to xlsx file"
    }

    fn run(
        &self,
        plugin: &ToXlsx,
        engine: &EngineInterface,
        call: &EvaluatedCall,
        input: PipelineData,
    ) -> Result<PipelineData, LabeledError> {
        let path_str = call.req::<String>(0)?;
        let mut path = PathBuf::from(&path_str);

        // If the path is relative, make it absolute using the current working directory
        if path.is_relative() {
            match engine.get_current_dir() {
                Ok(cwd) => {
                    path = PathBuf::from(cwd).join(path);
                }
                Err(err) => {
                    return Err(LabeledError {
                        msg: "Failed to get current working directory".into(),
                        labels: vec![],
                        code: None,
                        url: None,
                        help: Some(format!("Error: {}", err)),
                        inner: vec![],
                    });
                }
            }
        }

        let mut workbook = Workbook::new();
        let input_val = input.into_value(call.head)?;

        if let Err(err) = plugin.write_value(&mut workbook, "Sheet1", &input_val) {
            return Err(LabeledError {
                msg: "Failed to export data".into(),
                labels: vec![],
                code: None,
                url: None,
                help: Some(format!("Error: {}", err)),
                inner: vec![],
            });
        }

        if let Err(err) = workbook.save(path) {
            return Err(LabeledError {
                msg: "Failed to save file".into(),
                labels: vec![],
                code: None,
                url: None,
                help: Some(format!("Error: {}", err)),
                inner: vec![],
            });
        }

        Ok(PipelineData::empty())
    }
}

impl Plugin for ToXlsx {
    fn version(&self) -> String {
        env!("CARGO_PKG_VERSION").to_string()
    }

    fn commands(&self) -> Vec<Box<dyn PluginCommand<Plugin = Self>>> {
        vec![Box::new(ToXlsxCommand)]
    }
}

fn main() {
    serve_plugin(&mut ToXlsx::new(), nu_plugin::JsonSerializer);
}
