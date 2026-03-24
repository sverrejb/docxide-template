# Docxide Template - Type safe MS Word templates for Rust.

`docxide-template` is a Rust crate for working with MS Word templates such as `.docx`, `.docxm`, `.dotx`, `.dotm` It reads your template files, finds `{placeholder}` patterns in document text, and generates type-safe Rust structs with those placeholders as fields. The generated structs include a `save()` method that produces a new `.docx` with placeholders replaced by field values and a `to_bytes()` for outputting the raw bytes.

## Usage

```bash
cargo add docxide-template
```

Place your templates in a folder (e.g. `path/to/templates/`), containing `{PlaceholderName}` for variables.

Then invoke the macro:

```rust
use docxide_template::generate_templates;

generate_templates!("path/to/templates");

fn main() {
    // If templates/HelloWorld.docx contains {FirstName} and {Company}:
    let doc = HelloWorld {
        first_name: "Alice".into(),
        company: "Acme Corp".into(),
    };

    // Writes output/greeting.docx with placeholders replaced
    doc.save("output/greeting").unwrap();

    // Or outputs the filled template as bytes:
    doc.to_bytes()
}
```


Placeholders are converted to snake_case struct fields automatically:

| Placeholder in template | Struct field |
|------------------------|-------------|
| `{FirstName}` | `first_name` |
| `{last_name}` | `last_name` |
| `{middle-name}` | `middle_name` |
| `{companyName}` | `company_name` |
| `{USER_COUNTRY}` | `user_country` |
| `{first name}` | `first_name` |
| `{ ZipCode }` | `zip_code` |
| `{ZIPCODE}` | `zipcode` |

> Note: all upper- or lower-caps without a separator (like `ZIPCODE`) can't be split into words — use `ZIP_CODE` or another format if you want it to become `zip_code`.



## Polymorphism

All generated structs implement the `DocxTemplate` trait, which lets you write functions that accept any template type. This is useful for batch processing, pipelines, or anywhere you don't want to care about which specific template you're working with:

```rust
use docxide_template::{generate_templates, DocxTemplate, TemplateError};
use std::path::Path;

generate_templates!("templates");

fn export(template: &dyn DocxTemplate, path: &Path) -> Result<(), TemplateError> {
    template.save(path)?;
    Ok(())
}

fn main() {
    let documents: Vec<Box<dyn DocxTemplate>> = vec![
        Box::new(HelloWorld::new("Alice", "Acme Corp")),
        Box::new(Invoice::new("INV-001", "2025-01-15", "1234.00")),
    ];

    for (i, doc) in documents.iter().enumerate() {
        let path = format!("output/doc_{i}.docx");
        export(doc.as_ref(), Path::new(&path)).unwrap();
    }
}
```

The trait provides `to_bytes()`, `save()`, `replacements()`, and `template_path()`, so generic code has full access to both output generation and introspection. See the [batch export example](examples/batch_export/src/main.rs).

## Embedded templates

By default, `generate_templates!` reads template files from disk at runtime. If you want a fully self-contained binary with no runtime file dependencies, enable the `embed` feature:

```bash
cargo add docxide-template --features embed
```

With `embed` enabled, template bytes are baked into the binary at compile time via `include_bytes!`. The same `generate_templates!` macro is used.

## Examples

See the [`examples/`](examples/) directory for details.

**Save to file** — fill a template and write it to disk:
```bash
cargo run -p save-to-file
```

**To bytes** — fill a template and get the `.docx` as `Vec<u8>` in memory, useful for piping into other processing steps:
```bash
cargo run -p to-bytes
```

**Embedded** — template bytes baked into the binary, no runtime file access needed:
```bash
cargo run -p embedded
```

**Batch export** — process different template types uniformly via `dyn DocxTemplate`:
```bash
cargo run -p batch-export
```

## How it works

1. The proc macro scans the given directory for template files at compile time
2. Each file becomes a struct named after the filename (PascalCase)
3. `{placeholder}` patterns become struct fields (snake_case)
4. `save()` opens the original template, replaces all placeholders in the XML, and writes a new `.docx`

## License

MIT
