extern crate proc_macro;
mod codegen;
mod docx_extract;
mod naming;
mod placeholders;

use docx_rs::read_docx;
use proc_macro::TokenStream;
use quote::quote;
use std::{
    collections::HashMap,
    fs,
    path::PathBuf,
};

use syn::{parse_str, LitStr};

use codegen::generate_struct;
use docx_extract::{
    collect_text_from_document_children, collect_text_from_footer_children,
    collect_text_from_header_children, extract_text_from_xml_tags, is_valid_docx_file,
    print_docxide_message,
};
use naming::derive_type_name_from_filename;
use placeholders::generate_struct_content;

/// Scans a directory for `.docx` (and similar) template files and generates a typed struct for each one.
///
/// Template paths are stored as relative paths (as written in the macro call), so at
/// runtime they resolve against the current working directory. This means you can ship
/// a binary alongside its template folder and it will work from any machine.
///
/// With the `embed` feature enabled, template bytes are baked into the binary via
/// `include_bytes!`, making it fully self-contained with no runtime file dependencies.
///
/// # Usage
///
/// ```rust,ignore
/// use docxide_template::generate_templates;
///
/// generate_templates!("path/to/templates");
/// ```
///
/// For each `.docx` file, this generates a struct with:
/// - A field for each `{placeholder}` found in the document text (converted to snake_case)
/// - `new()` constructor taking all field values as `impl Into<String>`
/// - `save(path)` to write a filled-in `.docx` to disk
/// - `to_bytes()` to get the filled-in `.docx` as `Vec<u8>`
#[proc_macro]
pub fn generate_templates(input: TokenStream) -> TokenStream {
    let embed = cfg!(feature = "embed");

    let lit: LitStr = syn::parse(input).expect("expected a string literal, e.g. generate_templates!(\"path/to/templates\")");
    let folder_path = lit.value();

    let paths = fs::read_dir(&folder_path).unwrap_or_else(|e| panic!("Failed to read template directory {:?}: {}", folder_path, e));
    let mut structs = Vec::new();
    let mut seen_type_names: HashMap<String, PathBuf> = HashMap::new();

    for path in paths {
        let path = path.expect("Failed to read path").path();

        if !is_valid_docx_file(&path) {
            print_docxide_message("Invalid template file, skipping.", &path);
            continue;
        }

        let type_name = match derive_type_name_from_filename(&path) {
            Ok(name) if parse_str::<syn::Ident>(&name).is_ok() => name,
            other => {
                let stem = path.file_stem().and_then(|s| s.to_str()).unwrap_or("");
                if stem.starts_with(|c: char| c.is_ascii_digit()) {
                    let attempted = other.unwrap_or_default();
                    print_docxide_message(
                        &format!(
                            "Filename starts with a digit, which produces an invalid Rust type name `{}`. Skipping.",
                            if attempted.is_empty() { stem.to_string() } else { attempted }
                        ),
                        &path,
                    );
                } else {
                    print_docxide_message(
                        "Unable to derive a valid Rust type name from file name. Skipping.",
                        &path,
                    );
                }
                continue;
            }
        };

        if let Some(existing_path) = seen_type_names.get(&type_name) {
            panic!(
                "\n\n[Docxide-template] Type name collision: both {:?} and {:?} produce the struct name `{}`.\n\
                Rename one of the files to avoid this conflict.\n",
                existing_path, path, type_name
            );
        }
        seen_type_names.insert(type_name.clone(), path.clone());

        let type_ident = syn::Ident::new(type_name.as_str(), proc_macro::Span::call_site().into());

        let buf = match fs::read(&path) {
            Ok(buf) => buf,
            Err(_) => {
                print_docxide_message("Unable to read file content. Skipping.", &path);
                continue;
            }
        };

        let doc = match read_docx(&buf) {
            Ok(doc) => doc,
            Err(_) => {
                print_docxide_message("Unable to read docx content. Skipping.", &path);
                continue;
            }
        };

        let mut corpus = collect_text_from_document_children(doc.document.children);

        let section = &doc.document.section_property;
        for (_, header) in section.get_headers() {
            corpus.extend(collect_text_from_header_children(&header.children));
        }
        for (_, footer) in section.get_footers() {
            corpus.extend(collect_text_from_footer_children(&footer.children));
        }

        corpus.extend(extract_text_from_xml_tags(&buf, &["<a:t", "<m:t"]));

        let content = generate_struct_content(corpus);

        let abs_path = path.canonicalize().expect("Failed to canonicalize template path");
        let abs_path_str = abs_path.to_str().expect("Failed to convert path to string");
        let rel_path_str = path.to_str().expect("Failed to convert relative path to string");

        let template_struct = generate_struct(
            type_ident,
            abs_path_str,
            rel_path_str,
            &content.fields,
            &content.replacement_placeholders,
            &content.replacement_fields,
            embed,
        );

        structs.push(template_struct)
    }

    let combined = quote! {
        #(#structs)*
    };

    combined.into()
}
