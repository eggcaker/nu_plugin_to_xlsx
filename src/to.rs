use nu_plugin::{EvaluatedCall, LabeledError};
use nu_protocol::Value;
use rust_xlsxwriter::{ColNum, Format, RowNum, Workbook, Worksheet};


pub(crate) fn write_to_xlsx(
    _call: &EvaluatedCall,
    value: &Value,
    path: &str,
    sheet_name: &str,
) -> Result<(), LabeledError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let _ = worksheet.set_name(sheet_name);

    match &value {
        Value::List { vals, .. } => {
            for (row, val) in vals.iter().enumerate() {
                match val {
                    Value::Record { val, .. } => {
                        if row == 0 {
                            for (column, cn) in val.cols.iter().enumerate() {
                                let _ = worksheet.write_string(row as RowNum, column as ColNum, cn);
                            }
                        }

                        for (column, v) in val.vals.iter().enumerate() {
                            write_value(worksheet, (row + 1) as RowNum, column as ColNum, &v).unwrap();
                        }
                    }

                    _ => {}
                }
            }
        }
        _ => ()
    }

    let _ = workbook.save(path);

    Ok(())
}

fn write_value(
    worksheet: &mut Worksheet,
    row: RowNum,
    col: ColNum,
    value: &Value,
) -> Result<(), LabeledError> {
    let _ = match value {
        Value::String { val, .. } => worksheet.write_string(row, col, val),
        Value::Int { val, .. } => worksheet.write_number_with_format(row, col, *val as f64, &Format::new().set_num_format("###0")),
        Value::Float { val, .. } => worksheet.write_number(row, col, *val),
        Value::Date { val, .. } =>
             worksheet.write_string(row, col, &val.format("%s").to_string()),

        Value::Duration { val, .. } => worksheet.write_string(row, col, val.to_string()),
        Value::Bool { val, .. } => {
            worksheet.write_string(row, col, if *val { "true" } else { "false" })
        }
        Value::Filesize { val,..} => worksheet.write_string(row, col, val.to_string()),
        Value::Nothing { .. } => worksheet.write_blank(row, col, &Format::new()),
        _ => todo!(),
    };

    Ok(())
}
