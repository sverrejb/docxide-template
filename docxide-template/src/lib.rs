//! Type-safe `.docx` template engine.
//!
//! Use [`generate_templates!`] to scan a directory of `.docx` files at compile time
//! and generate a struct per template. See the [README](https://github.com/sverrejb/docxide-template)
//! for full usage instructions.

pub use docxide_template_derive::generate_templates;

use std::io::{Cursor, Read, Write};
use std::path::Path;

/// Error type returned by template `save()` and `to_bytes()` methods.
#[derive(Debug)]
pub enum TemplateError {
    /// An I/O error (reading template, writing output, creating directories).
    Io(std::io::Error),
    /// The `.docx` template is malformed (bad zip archive, invalid XML encoding).
    InvalidTemplate(String),
}

impl std::fmt::Display for TemplateError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(e) => write!(f, "{}", e),
            Self::InvalidTemplate(msg) => write!(f, "invalid template: {}", msg),
        }
    }
}

impl std::error::Error for TemplateError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(e) => Some(e),
            Self::InvalidTemplate(_) => None,
        }
    }
}

impl From<std::io::Error> for TemplateError {
    fn from(e: std::io::Error) -> Self { Self::Io(e) }
}

impl From<zip::result::ZipError> for TemplateError {
    fn from(e: zip::result::ZipError) -> Self {
        match e {
            zip::result::ZipError::Io(io_err) => Self::Io(io_err),
            other => Self::InvalidTemplate(other.to_string()),
        }
    }
}

impl From<std::string::FromUtf8Error> for TemplateError {
    fn from(e: std::string::FromUtf8Error) -> Self { Self::InvalidTemplate(e.to_string()) }
}

#[doc(hidden)]
pub trait DocxTemplate: __private::Sealed {
    fn template_path(&self) -> &Path;
    fn replacements(&self) -> Vec<(&str, &str)>;
}

#[doc(hidden)]
pub mod __private {
    use super::*;

    pub trait Sealed {}

    pub fn save_docx<T: DocxTemplate, P: AsRef<Path>>(
        template: &T,
        output_path: P,
    ) -> Result<(), TemplateError> {
        let template_bytes = std::fs::read(template.template_path())?;
        save_docx_bytes(&template_bytes, output_path.as_ref(), &template.replacements())
    }

    pub fn build_docx_bytes(
        template_bytes: &[u8],
        replacements: &[(&str, &str)],
    ) -> Result<Vec<u8>, TemplateError> {
        let cursor = Cursor::new(template_bytes);
        let mut archive = zip::read::ZipArchive::new(cursor)?;

        let mut output_buf = Cursor::new(Vec::new());
        let mut zip_writer = zip::write::ZipWriter::new(&mut output_buf);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .compression_level(Some(6));

        for i in 0..archive.len() {
            let mut file = archive.by_index(i)?;
            let file_name = file.name().to_string();

            let mut contents = Vec::new();
            file.read_to_end(&mut contents)?;

            if file_name.ends_with(".xml") || file_name.ends_with(".rels") {
                let xml = String::from_utf8(contents)?;
                let replaced = replace_placeholders_in_xml(&xml, replacements);
                contents = replaced.into_bytes();
            }

            zip_writer.start_file(&file_name, options)?;
            zip_writer.write_all(&contents)?;
        }

        zip_writer.finish()?;
        Ok(output_buf.into_inner())
    }

    pub fn save_docx_bytes(
        template_bytes: &[u8],
        output_path: &Path,
        replacements: &[(&str, &str)],
    ) -> Result<(), TemplateError> {
        let bytes = build_docx_bytes(template_bytes, replacements)?;
        if let Some(parent) = output_path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        std::fs::write(output_path, bytes)?;
        Ok(())
    }
}

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn replace_placeholders_in_xml(xml: &str, replacements: &[(&str, &str)]) -> String {
    let mut text_spans: Vec<(usize, usize, String)> = Vec::new();
    let mut search_start = 0;
    while let Some(tag_start) = xml[search_start..].find("<w:t") {
        let tag_start = search_start + tag_start;
        let content_start = match xml[tag_start..].find('>') {
            Some(pos) => tag_start + pos + 1,
            None => break,
        };
        let content_end = match xml[content_start..].find("</w:t>") {
            Some(pos) => content_start + pos,
            None => break,
        };
        let text = xml[content_start..content_end].to_string();
        text_spans.push((content_start, content_end, text));
        search_start = content_end + 6;
    }

    if text_spans.is_empty() {
        return xml.to_string();
    }

    let concatenated: String = text_spans.iter().map(|(_, _, t)| t.as_str()).collect();

    let offset_map: Vec<(usize, usize)> = text_spans
        .iter()
        .enumerate()
        .flat_map(|(span_idx, (_, _, text))| {
            (0..text.len()).map(move |char_offset| (span_idx, char_offset))
        })
        .collect();

    let mut span_replacements: Vec<Vec<(usize, usize, String)>> = vec![Vec::new(); text_spans.len()];
    for &(placeholder, value) in replacements {
        let mut start = 0;
        while let Some(found) = concatenated[start..].find(placeholder) {
            let match_start = start + found;
            let match_end = match_start + placeholder.len();
            if match_start >= offset_map.len() || match_end > offset_map.len() {
                break;
            }

            let (start_span, start_off) = offset_map[match_start];
            let (end_span, _) = offset_map[match_end - 1];
            let end_off_exclusive = offset_map[match_end - 1].1 + 1;

            if start_span == end_span {
                span_replacements[start_span].push((start_off, end_off_exclusive, escape_xml(value)));
            } else {
                let first_span_text = &text_spans[start_span].2;
                span_replacements[start_span].push((start_off, first_span_text.len(), escape_xml(value)));
                for mid in (start_span + 1)..end_span {
                    let mid_len = text_spans[mid].2.len();
                    span_replacements[mid].push((0, mid_len, String::new()));
                }
                span_replacements[end_span].push((0, end_off_exclusive, String::new()));
            }
            start = match_end;
        }
    }

    let mut result = xml.to_string();
    for (span_idx, (content_start, content_end, _)) in text_spans.iter().enumerate().rev() {
        let mut span_text = result[*content_start..*content_end].to_string();
        let mut reps = span_replacements[span_idx].clone();
        reps.sort_by(|a, b| b.0.cmp(&a.0));
        for (from, to, replacement) in reps {
            let safe_to = to.min(span_text.len());
            span_text = format!("{}{}{}", &span_text[..from], replacement, &span_text[safe_to..]);
        }
        result = format!("{}{}{}", &result[..*content_start], span_text, &result[*content_end..]);
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn replace_single_run_placeholder() {
        let xml = r#"<w:t>{Name}</w:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert_eq!(result, r#"<w:t>Alice</w:t>"#);
    }

    #[test]
    fn replace_placeholder_split_across_runs() {
        let xml = r#"<w:t>{Na</w:t><w:t>me}</w:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert_eq!(result, r#"<w:t>Alice</w:t><w:t></w:t>"#);
    }

    #[test]
    fn replace_placeholder_with_inner_whitespace() {
        let xml = r#"<w:t>Hello { Name }!</w:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{ Name }", "Alice")]);
        assert_eq!(result, r#"<w:t>Hello Alice!</w:t>"#);
    }

    #[test]
    fn replace_both_whitespace_variants() {
        let xml = r#"<w:t>{Name} and { Name }</w:t>"#;
        let result = replace_placeholders_in_xml(
            xml,
            &[("{Name}", "Alice"), ("{ Name }", "Alice")],
        );
        assert_eq!(result, r#"<w:t>Alice and Alice</w:t>"#);
    }

    #[test]
    fn replace_multiple_placeholders() {
        let xml = r#"<w:t>Hello {First} {Last}!</w:t>"#;
        let result = replace_placeholders_in_xml(
            xml,
            &[("{First}", "Alice"), ("{Last}", "Smith")],
        );
        assert_eq!(result, r#"<w:t>Hello Alice Smith!</w:t>"#);
    }

    #[test]
    fn no_placeholders_returns_unchanged() {
        let xml = r#"<w:t>No placeholders here</w:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert_eq!(result, xml);
    }

    #[test]
    fn no_wt_tags_returns_unchanged() {
        let xml = r#"<w:p>plain paragraph</w:p>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert_eq!(result, xml);
    }

    #[test]
    fn empty_replacements_returns_unchanged() {
        let xml = r#"<w:t>{Name}</w:t>"#;
        let result = replace_placeholders_in_xml(xml, &[]);
        assert_eq!(result, xml);
    }

    #[test]
    fn preserves_wt_attributes() {
        let xml = r#"<w:t xml:space="preserve">{Name}</w:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert_eq!(result, r#"<w:t xml:space="preserve">Alice</w:t>"#);
    }

    #[test]
    fn replace_whitespace_placeholder_split_across_runs() {
        // Mimics Word splitting "{ foo }" across 5 <w:t> tags
        let xml = r#"<w:t>{</w:t><w:t xml:space="preserve"> </w:t><w:t>foo</w:t><w:t xml:space="preserve"> </w:t><w:t>}</w:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{ foo }", "bar")]);
        assert!(
            !result.contains("foo"),
            "placeholder not replaced: {}",
            result
        );
        assert!(result.contains("bar"), "value not present: {}", result);
    }

    #[test]
    fn replace_whitespace_placeholder_with_prooferr_between_runs() {
        // Exact XML from Word: proofErr tag sits between <w:t> runs
        let xml = concat!(
            r#"<w:r><w:t>{foo}</w:t></w:r>"#,
            r#"<w:r><w:t>{</w:t></w:r>"#,
            r#"<w:r><w:t xml:space="preserve"> </w:t></w:r>"#,
            r#"<w:r><w:t>foo</w:t></w:r>"#,
            r#"<w:proofErr w:type="gramEnd"/>"#,
            r#"<w:r><w:t xml:space="preserve"> </w:t></w:r>"#,
            r#"<w:r><w:t>}</w:t></w:r>"#,
        );
        let result = replace_placeholders_in_xml(
            xml,
            &[("{foo}", "bar"), ("{ foo }", "bar")],
        );
        // Both {foo} and { foo } should be replaced
        assert!(
            !result.contains("foo"),
            "placeholder not replaced: {}",
            result
        );
    }

    #[test]
    fn replace_all_variants_in_full_document() {
        // Mimics HeadFootTest.docx: {header} x2, {foo}, { foo } split, {  foo  } split
        let xml = concat!(
            r#"<w:t>{header}</w:t>"#,
            r#"<w:t>{header}</w:t>"#,
            r#"<w:t>{foo}</w:t>"#,
            // { foo } split across 5 runs
            r#"<w:t>{</w:t>"#,
            r#"<w:t xml:space="preserve"> </w:t>"#,
            r#"<w:t>foo</w:t>"#,
            r#"<w:t xml:space="preserve"> </w:t>"#,
            r#"<w:t>}</w:t>"#,
            // {  foo  } split across 6 runs
            r#"<w:t>{</w:t>"#,
            r#"<w:t xml:space="preserve"> </w:t>"#,
            r#"<w:t xml:space="preserve"> </w:t>"#,
            r#"<w:t>foo</w:t>"#,
            r#"<w:t xml:space="preserve">  </w:t>"#,
            r#"<w:t>}</w:t>"#,
        );
        let result = replace_placeholders_in_xml(
            xml,
            &[
                ("{header}", "TITLE"),
                ("{foo}", "BAR"),
                ("{ foo }", "BAR"),
                ("{  foo  }", "BAR"),
            ],
        );
        assert!(
            !result.contains("header"),
            "{{header}} not replaced: {}",
            result,
        );
        assert!(
            !result.contains("foo"),
            "foo variant not replaced: {}",
            result,
        );
    }

    #[test]
    fn duplicate_replacement_does_not_break_later_spans() {
        // {header} appears twice in replacements
        let xml = concat!(
            r#"<w:t>{header}</w:t>"#,
            r#"<w:t>{header}</w:t>"#,
            r#"<w:t>{foo}</w:t>"#,
            r#"<w:t>{</w:t>"#,
            r#"<w:t xml:space="preserve"> </w:t>"#,
            r#"<w:t>foo</w:t>"#,
            r#"<w:t xml:space="preserve"> </w:t>"#,
            r#"<w:t>}</w:t>"#,
        );
        let result = replace_placeholders_in_xml(
            xml,
            &[
                // duplicate {header}
                ("{header}", "TITLE"),
                ("{header}", "TITLE"),
                ("{foo}", "BAR"),
                ("{ foo }", "BAR"),
            ],
        );
        // Check if { foo } was replaced despite the duplicate
        assert!(
            !result.contains("foo"),
            "foo not replaced when duplicate header present: {}",
            result,
        );
    }

    #[test]
    fn replace_headfoottest_template() {
        let template_path = Path::new("../test-crate/templates/HeadFootTest.docx");
        if !template_path.exists() {
            return;
        }
        let template_bytes = std::fs::read(template_path).unwrap();
        let result = __private::build_docx_bytes(
            &template_bytes,
            &[
                ("{header}", "TITLE"),
                ("{foo}", "BAR"),
                ("{ foo }", "BAR"),
                ("{  foo  }", "BAR"),
                ("{top}", "TOP"),
                ("{bottom}", "BOT"),
            ],
        )
        .unwrap();

        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut doc_xml = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut doc_xml)
            .unwrap();

        assert!(!doc_xml.contains("{header}"), "header placeholder not replaced");
        assert!(!doc_xml.contains("{foo}"), "foo placeholder not replaced");
        assert!(!doc_xml.contains("{ foo }"), "spaced foo placeholder not replaced");
    }

    #[test]
    fn build_docx_bytes_produces_valid_zip() {
        let template_path = Path::new("../test-crate/templates/HelloWorld.docx");
        if !template_path.exists() {
            return;
        }
        let template_bytes = std::fs::read(template_path).unwrap();
        let result = __private::build_docx_bytes(
            &template_bytes,
            &[("{ firstName }", "Test"), ("{ productName }", "Lib")],
        )
        .unwrap();

        assert!(!result.is_empty());
        let cursor = Cursor::new(&result);
        let archive = zip::ZipArchive::new(cursor).expect("output should be a valid zip");
        assert!(archive.len() > 0);
    }

    #[test]
    fn escape_xml_special_characters() {
        let xml = r#"<w:t>{Name}</w:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice & Bob")]);
        assert_eq!(result, r#"<w:t>Alice &amp; Bob</w:t>"#);

        let result = replace_placeholders_in_xml(xml, &[("{Name}", "<script>")]);
        assert_eq!(result, r#"<w:t>&lt;script&gt;</w:t>"#);

        let result = replace_placeholders_in_xml(xml, &[("{Name}", "a < b & c > d")]);
        assert_eq!(result, r#"<w:t>a &lt; b &amp; c &gt; d</w:t>"#);

        let result = replace_placeholders_in_xml(xml, &[("{Name}", r#"She said "hello""#)]);
        assert_eq!(result, r#"<w:t>She said &quot;hello&quot;</w:t>"#);

        let result = replace_placeholders_in_xml(xml, &[("{Name}", "it's")]);
        assert_eq!(result, r#"<w:t>it&apos;s</w:t>"#);
    }

    #[test]
    fn escape_xml_split_across_runs() {
        let xml = r#"<w:t>{Na</w:t><w:t>me}</w:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "A&B")]);
        assert_eq!(result, r#"<w:t>A&amp;B</w:t><w:t></w:t>"#);
    }

    #[test]
    fn escape_xml_in_headfoottest_template() {
        let template_path = Path::new("../test-crate/templates/HeadFootTest.docx");
        if !template_path.exists() {
            return;
        }
        let template_bytes = std::fs::read(template_path).unwrap();
        let result = __private::build_docx_bytes(
            &template_bytes,
            &[
                ("{header}", "Tom & Jerry"),
                ("{foo}", "x < y"),
                ("{ foo }", "x < y"),
                ("{  foo  }", "x < y"),
                ("{top}", "A > B"),
                ("{bottom}", "C & D"),
            ],
        )
        .unwrap();

        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut doc_xml = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut doc_xml)
            .unwrap();

        assert!(!doc_xml.contains("Tom & Jerry"), "raw ampersand should be escaped");
        assert!(doc_xml.contains("Tom &amp; Jerry"), "escaped value should be present");
        assert!(!doc_xml.contains("x < y"), "raw less-than should be escaped");
    }

    #[test]
    fn build_docx_bytes_replaces_content() {
        let template_path = Path::new("../test-crate/templates/HelloWorld.docx");
        if !template_path.exists() {
            return;
        }
        let template_bytes = std::fs::read(template_path).unwrap();
        let result = __private::build_docx_bytes(
            &template_bytes,
            &[("{ firstName }", "Alice"), ("{ productName }", "Docxide")],
        )
        .unwrap();

        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut doc_xml = String::new();
        archive
            .by_name("word/document.xml")
            .unwrap()
            .read_to_string(&mut doc_xml)
            .unwrap();
        assert!(doc_xml.contains("Alice"));
        assert!(doc_xml.contains("Docxide"));
        assert!(!doc_xml.contains("firstName"));
        assert!(!doc_xml.contains("productName"));
    }
}
