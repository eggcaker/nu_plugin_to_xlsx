use nu_plugin::LabeledError;
use nu_protocol::{Span, Value};
use rust_xlsxwriter::{ColNum, Format, RowNum, Workbook, Worksheet};

pub(crate) fn write_to_xlsx(
    _document: &Value,
    path: &str,
    sheet_name: &str,
) -> Result<(), LabeledError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    let _ = worksheet.set_name(sheet_name);

    let val = Value::String {
        val: String::from("hello"),
        internal_span: Span::unknown(),
    };
    let world = Value::String {
        val: String::from("world"),
        internal_span: Span::unknown(),
    };
    write_value(worksheet, 1, 1, &val).unwrap();
    write_value(worksheet, 2, 2, &world).unwrap();
    /*
    let Value::Record { cols, .. } = document else { todo!() };

    for (i, col) in cols.iter().enumerate() {
        write_value(&mut worksheet, 0, i as u16, Value::String { val: col.clone(), span: Span::unknown() }).unwrap();
        let data = &document.get_data_by_key(col).unwrap();
        for (j, val) in data.iter().enumerate() {
            write_value(&mut worksheet, (j+1) as RowNum, i as ColNum, val).unwrap();
        }
    }
    */

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
        Value::Int { val, .. } => worksheet.write_number(row, col, *val as f64),
        Value::Float { val, .. } => worksheet.write_number(row, col, *val),
        Value::Bool { val, .. } => {
            worksheet.write_string(row, col, if *val { "true" } else { "false" })
        }
        Value::Nothing { .. } => worksheet.write_blank(row, col, &Format::new()),
        _ => todo!(),
    };

    Ok(())
}
