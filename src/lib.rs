mod to;

use nu_plugin::{EvaluatedCall, LabeledError, Plugin};
use nu_protocol::{Category, PluginExample, PluginSignature, SyntaxShape, Type, Value};

pub struct XLSX;

impl XLSX {
    pub fn to(&self, call: &EvaluatedCall, input: &Value) -> Result<Value, LabeledError> {
        let path: String = call.req(0)?;

        let sheet_name = call.opt(1)?.unwrap_or("Sheet1".to_string());

        to::write_to_xlsx(input, &path, &sheet_name)?;

        let val = Value::String {
            val: String::from(""),
            internal_span: call.head,
        };
        return Ok(val);
    }
}

impl Plugin for XLSX {
    fn signature(&self) -> Vec<PluginSignature> {
        vec![PluginSignature::build("to xlsx")
            .usage("Convert table to excel(.xlsx) file.")
            .required("file_path", SyntaxShape::String, "required xlsx file path")
            .optional(
                "sheet_name",
                SyntaxShape::String,
                "Sheet name for the table(default Sheet1)",
            )
            .input_output_type(Type::Any, Type::Nothing)
            .plugin_examples(vec![PluginExample {
                example: "echo [[name]; [bob]] | to xlsx".into(),
                description: "Convert a table or record and save to xlsx file".into(),
                result: None,
            }])
            .category(Category::Experimental)]
    }

    fn run(
        &mut self,
        name: &str,
        call: &EvaluatedCall,
        input: &Value,
    ) -> Result<Value, LabeledError> {
        match name {
            "to xlsx" => Ok( self.to(call, input)?),
            _ => Err(LabeledError{
               label: "Plugin call with wrong name signature".into(),
               msg: "the signature used to call the plugin does not match any name in the plugin signature vector".into(),
               span:  Some(call.head),
            }),
        }
    }
}
