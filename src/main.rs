use nu_plugin::{serve_plugin, Plugin, PluginCommand, EvaluatedCall, EngineInterface};
use nu_protocol::{Category, LabeledError, PipelineData, Signature, SyntaxShape, Type, Value};
use rust_xlsxwriter::Workbook;
use std::path::PathBuf;

struct ToXlsx;

impl ToXlsx {
    fn new() -> Self {
        Self
    }

    fn write_value(&self, workbook: &mut Workbook, worksheet_name: &str, value: &Value) -> Result<(), Box<dyn std::error::Error>> {
        let mut worksheet = workbook.add_worksheet().set_name(worksheet_name)?;

        match value {
            Value::List { vals, .. } => {
                // Handle list of records (table-like data)
                if let Some(Value::Record { val: first_record, .. }) = vals.first() {
                    // Write headers
                    let headers: Vec<String> = first_record.columns().into_iter().map(|s| s.to_string()).collect();
                    for (col, header) in headers.iter().enumerate() {
                        worksheet.write_string(0, col as u16, header)?;
                    }

                    // Write data
                    for (row, record_value) in vals.iter().enumerate() {
                        if let Value::Record { val, .. } = record_value {
                            for (col, header) in headers.iter().enumerate() {
                                if let Some(cell_value) = val.get(header) {
                                    self.write_cell(&mut worksheet, (row + 1) as u32, col as u16, cell_value)?;
                                }
                            }
                        }
                    }
                } else {
                    // Handle simple list
                    for (row, item) in vals.iter().enumerate() {
                        self.write_cell(&mut worksheet, row as u32, 0, item)?;
                    }
                }
            }
            Value::Record { val, .. } => {
                // Write record as two columns: key and value
                worksheet.write_string(0, 0, "Key")?;
                worksheet.write_string(0, 1, "Value")?;

                for (row, (key, value)) in val.iter().enumerate() {
                    worksheet.write_string((row + 1) as u32, 0, key)?;
                    self.write_cell(&mut worksheet, (row + 1) as u32, 1, value)?;
                }
            }
            _ => {
                // Write single value
                self.write_cell(&mut worksheet, 0, 0, value)?;
            }
        }

        Ok(())
    }

    fn write_cell(&self, worksheet: &mut rust_xlsxwriter::Worksheet, row: u32, col: u16, value: &Value) -> Result<(), Box<dyn std::error::Error>> {
        match value {
            Value::String { val, .. } => worksheet.write_string(row, col, val)?,
            Value::Int { val, .. } => worksheet.write_number(row, col, *val as f64)?,
            Value::Float { val, .. } => worksheet.write_number(row, col, *val)?,
            Value::Bool { val, .. } => worksheet.write_boolean(row, col, *val)?,
            _ => worksheet.write_string(row, col, &format!("{:?}", value))?,
        };
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
