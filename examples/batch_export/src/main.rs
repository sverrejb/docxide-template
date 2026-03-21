use docxide_template::{generate_templates, DocxTemplate, TemplateError};
use std::path::Path;

generate_templates!("examples/batch_export/templates");

/// Writes any template to disk and prints the output path.
fn export(template: &dyn DocxTemplate, path: &Path) -> Result<(), TemplateError> {
    template.save(path)?;
    println!("  Exported {}", path.display());
    Ok(())
}

fn main() -> Result<(), TemplateError> {
    let out = Path::new("examples/batch_export/output");

    // Different template types can be collected into one Vec and processed uniformly.
    let documents: Vec<(&str, Box<dyn DocxTemplate>)> = vec![
        ("alice_greeting", Box::new(HelloWorld::new("Alice", "docxide"))),
        ("bob_greeting",   Box::new(HelloWorld::new("Bob", "the template team"))),
        ("alice_table",    Box::new(TablePlaceholders::new("Alice", "Oslo"))),
        ("bob_table",      Box::new(TablePlaceholders::new("Bob", "Bergen"))),
    ];

    println!("Exporting {} documents:", documents.len());
    for (name, doc) in &documents {
        export(doc.as_ref(), &out.join(name).with_extension("docx"))?;
    }

    Ok(())
}
