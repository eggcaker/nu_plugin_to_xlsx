mod to; 


use nu_plugin::{EvaluatedCall, LabeledError, Plugin};
use nu_protocol::{Category, PluginExample, PluginSignature, Type, Value};


pub struct XLSX;

impl XLSX {
    pub fn to (&self, call:&EvaluatedCall, input: &Value) -> Result<Value, LabeledError> {
        let path = match call.positional.get(0) {
            Some(v) => match v.as_string() {
            Ok(s) => s,
            _ => return Err(LabeledError{
               label:"Expected a path".into(),
               msg: "Expected a path".into(),
               span: Some(call.head),
            }),
        },
            None => "UnTittled.xlsx".to_string()
        };

        to::write_to_xlsx(input, &path)?;

        let val = Value::String { val: String::from(""), internal_span: call.head };
        return Ok(val);

       }
} 

impl Plugin for XLSX {
    fn signature(&self) -> Vec<PluginSignature> {
        vec![
            PluginSignature::build("to xlsx")
            .usage("Convert table to excel(.xlsx) file.")
            .input_output_type(Type::Record(vec![]), Type::Nothing)
            .plugin_examples(vec![PluginExample {
                example: "echo [[name]; [bob]] | to xlsx".into(),
                description: "Converts a table to xlsx".into(),
                result: None,
            }])
            .category(Category::Experimental)
        ]
    }


    fn run (
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