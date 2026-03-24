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

/// Trait implemented by all generated template structs.
///
/// Enables polymorphic use of templates via `&dyn DocxTemplate` or generics:
///
/// ```ignore
/// use docxide_template::DocxTemplate;
///
/// fn process(template: &dyn DocxTemplate) -> Result<Vec<u8>, docxide_template::TemplateError> {
///     template.to_bytes()
/// }
/// ```
pub trait DocxTemplate {
    /// Returns the filesystem path to the original `.docx` template.
    fn template_path(&self) -> &Path;

    /// Returns the list of `(placeholder, value)` pairs for substitution.
    fn replacements(&self) -> Vec<(&str, &str)>;

    /// Produces the filled-in `.docx` as an in-memory byte vector.
    fn to_bytes(&self) -> Result<Vec<u8>, TemplateError>;

    /// Writes the filled-in `.docx` to the given path.
    ///
    /// Creates parent directories if they do not exist.
    /// The path is used as-is — callers should include the `.docx` extension.
    fn save(&self, path: &Path) -> Result<(), TemplateError> {
        let bytes = self.to_bytes()?;
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)?;
        }
        std::fs::write(path, bytes)?;
        Ok(())
    }
}

#[doc(hidden)]
pub mod __private {
    use super::*;

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
                let mut replaced = replace_placeholders_in_xml(&xml, replacements);
                if file_name == "[Content_Types].xml" {
                    replaced = convert_template_content_types(&replaced);
                }
                contents = replaced.into_bytes();
            }

            zip_writer.start_file(&file_name, options)?;
            zip_writer.write_all(&contents)?;
        }

        zip_writer.finish()?;
        Ok(output_buf.into_inner())
    }
}

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn replace_for_tag(xml: &str, replacements: &[(&str, &str)], open_prefix: &str, close_tag: &str) -> String {
    let mut text_spans: Vec<(usize, usize, String)> = Vec::new();
    let mut search_start = 0;
    while let Some(tag_start) = xml[search_start..].find(open_prefix) {
        let tag_start = search_start + tag_start;
        let after_prefix = tag_start + open_prefix.len();
        if after_prefix < xml.len() && !matches!(xml.as_bytes()[after_prefix], b'>' | b' ') {
            search_start = after_prefix;
            continue;
        }
        let content_start = match xml[tag_start..].find('>') {
            Some(pos) => tag_start + pos + 1,
            None => break,
        };
        let content_end = match xml[content_start..].find(close_tag) {
            Some(pos) => content_start + pos,
            None => break,
        };
        let text = xml[content_start..content_end].to_string();
        text_spans.push((content_start, content_end, text));
        search_start = content_end + close_tag.len();
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

fn replace_placeholders_in_xml(xml: &str, replacements: &[(&str, &str)]) -> String {
    let result = replace_for_tag(xml, replacements, "<w:t", "</w:t>");
    let result = replace_for_tag(&result, replacements, "<a:t", "</a:t>");
    replace_for_tag(&result, replacements, "<m:t", "</m:t>")
}

//we do this to also support dotx/dotm files
fn convert_template_content_types(xml: &str) -> String {
    xml.replace(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    )
    .replace(
        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
    )
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
    fn replace_in_table_cell_xml() {
        let xml = concat!(
            r#"<w:tbl><w:tr><w:tc>"#,
            r#"<w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>"#,
            r#"<w:p><w:r><w:t>{Name}</w:t></w:r></w:p>"#,
            r#"</w:tc></w:tr></w:tbl>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert!(result.contains("Alice"), "placeholder in table cell not replaced: {}", result);
        assert!(!result.contains("{Name}"), "placeholder still present: {}", result);
    }

    #[test]
    fn replace_in_nested_table_xml() {
        let xml = concat!(
            r#"<w:tbl><w:tr><w:tc>"#,
            r#"<w:tbl><w:tr><w:tc>"#,
            r#"<w:p><w:r><w:t>{Inner}</w:t></w:r></w:p>"#,
            r#"</w:tc></w:tr></w:tbl>"#,
            r#"</w:tc></w:tr></w:tbl>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{Inner}", "Nested")]);
        assert!(result.contains("Nested"), "placeholder in nested table not replaced: {}", result);
        assert!(!result.contains("{Inner}"), "placeholder still present: {}", result);
    }

    #[test]
    fn replace_multiple_cells_same_row() {
        let xml = concat!(
            r#"<w:tbl><w:tr>"#,
            r#"<w:tc><w:p><w:r><w:t>{First}</w:t></w:r></w:p></w:tc>"#,
            r#"<w:tc><w:p><w:r><w:t>{Last}</w:t></w:r></w:p></w:tc>"#,
            r#"<w:tc><w:p><w:r><w:t>{Age}</w:t></w:r></w:p></w:tc>"#,
            r#"</w:tr></w:tbl>"#,
        );
        let result = replace_placeholders_in_xml(
            xml,
            &[("{First}", "Alice"), ("{Last}", "Smith"), ("{Age}", "30")],
        );
        assert!(result.contains("Alice"), "First not replaced: {}", result);
        assert!(result.contains("Smith"), "Last not replaced: {}", result);
        assert!(result.contains("30"), "Age not replaced: {}", result);
        assert!(!result.contains("{First}") && !result.contains("{Last}") && !result.contains("{Age}"),
            "placeholders still present: {}", result);
    }

    #[test]
    fn replace_in_footnote_xml() {
        let xml = concat!(
            r#"<w:footnotes>"#,
            r#"<w:footnote w:type="normal" w:id="1">"#,
            r#"<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>"#,
            r#"<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>"#,
            r#"<w:r><w:t xml:space="preserve"> </w:t></w:r>"#,
            r#"<w:r><w:t>{Source}</w:t></w:r>"#,
            r#"</w:p>"#,
            r#"</w:footnote>"#,
            r#"</w:footnotes>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{Source}", "Wikipedia")]);
        assert!(result.contains("Wikipedia"), "placeholder in footnote not replaced: {}", result);
        assert!(!result.contains("{Source}"), "placeholder still present: {}", result);
    }

    #[test]
    fn replace_in_endnote_xml() {
        let xml = concat!(
            r#"<w:endnotes>"#,
            r#"<w:endnote w:type="normal" w:id="1">"#,
            r#"<w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr>"#,
            r#"<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r>"#,
            r#"<w:r><w:t xml:space="preserve"> </w:t></w:r>"#,
            r#"<w:r><w:t>{Citation}</w:t></w:r>"#,
            r#"</w:p>"#,
            r#"</w:endnote>"#,
            r#"</w:endnotes>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{Citation}", "Doe, 2024")]);
        assert!(result.contains("Doe, 2024"), "placeholder in endnote not replaced: {}", result);
        assert!(!result.contains("{Citation}"), "placeholder still present: {}", result);
    }

    #[test]
    fn replace_in_comment_xml() {
        let xml = concat!(
            r#"<w:comments>"#,
            r#"<w:comment w:id="0" w:author="Author" w:date="2024-01-01T00:00:00Z">"#,
            r#"<w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr>"#,
            r#"<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>"#,
            r#"<w:r><w:t>{ReviewNote}</w:t></w:r>"#,
            r#"</w:p>"#,
            r#"</w:comment>"#,
            r#"</w:comments>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{ReviewNote}", "Approved")]);
        assert!(result.contains("Approved"), "placeholder in comment not replaced: {}", result);
        assert!(!result.contains("{ReviewNote}"), "placeholder still present: {}", result);
    }

    #[test]
    fn replace_in_sdt_xml() {
        let xml = concat!(
            r#"<w:sdt>"#,
            r#"<w:sdtPr><w:alias w:val="Title"/></w:sdtPr>"#,
            r#"<w:sdtContent>"#,
            r#"<w:p><w:r><w:t>{Title}</w:t></w:r></w:p>"#,
            r#"</w:sdtContent>"#,
            r#"</w:sdt>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{Title}", "Report")]);
        assert!(result.contains("Report"), "placeholder in sdt not replaced: {}", result);
        assert!(!result.contains("{Title}"), "placeholder still present: {}", result);
    }

    #[test]
    fn replace_in_hyperlink_display_text() {
        let xml = concat!(
            r#"<w:p>"#,
            r#"<w:hyperlink r:id="rId5" w:history="1">"#,
            r#"<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>"#,
            r#"<w:t>{LinkText}</w:t></w:r>"#,
            r#"</w:hyperlink>"#,
            r#"</w:p>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{LinkText}", "Click here")]);
        assert!(result.contains("Click here"), "placeholder in hyperlink not replaced: {}", result);
        assert!(!result.contains("{LinkText}"), "placeholder still present: {}", result);
    }

    #[test]
    fn replace_in_textbox_xml() {
        let xml = concat!(
            r#"<wps:txbx>"#,
            r#"<w:txbxContent>"#,
            r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr>"#,
            r#"<w:r><w:rPr><w:b/></w:rPr><w:t>{BoxTitle}</w:t></w:r>"#,
            r#"</w:p>"#,
            r#"</w:txbxContent>"#,
            r#"</wps:txbx>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{BoxTitle}", "Important")]);
        assert!(result.contains("Important"), "placeholder in textbox not replaced: {}", result);
        assert!(!result.contains("{BoxTitle}"), "placeholder still present: {}", result);
    }

    #[test]
    fn replace_placeholder_split_across_three_runs() {
        let xml = concat!(
            r#"<w:r><w:t>{pl</w:t></w:r>"#,
            r#"<w:r><w:t>ace</w:t></w:r>"#,
            r#"<w:r><w:t>holder}</w:t></w:r>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{placeholder}", "value")]);
        assert!(result.contains("value"), "placeholder split across 3 runs not replaced: {}", result);
        assert!(!result.contains("{pl"), "leftover fragment: {}", result);
        assert!(!result.contains("holder}"), "leftover fragment: {}", result);
    }

    #[test]
    fn replace_placeholder_split_across_four_runs() {
        let xml = concat!(
            r#"<w:r><w:t>{p</w:t></w:r>"#,
            r#"<w:r><w:t>la</w:t></w:r>"#,
            r#"<w:r><w:t>ceh</w:t></w:r>"#,
            r#"<w:r><w:t>older}</w:t></w:r>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{placeholder}", "value")]);
        assert!(result.contains("value"), "placeholder split across 4 runs not replaced: {}", result);
        assert!(!result.contains("placeholder"), "leftover fragment: {}", result);
    }

    #[test]
    fn replace_adjacent_placeholders_no_space() {
        let xml = r#"<w:r><w:t>{first}{last}</w:t></w:r>"#;
        let result = replace_placeholders_in_xml(xml, &[("{first}", "Alice"), ("{last}", "Smith")]);
        assert_eq!(result, r#"<w:r><w:t>AliceSmith</w:t></w:r>"#);
    }

    #[test]
    fn replace_with_bookmark_markers_between_runs() {
        let xml = concat!(
            r#"<w:r><w:t>{Na</w:t></w:r>"#,
            r#"<w:bookmarkStart w:id="0" w:name="bookmark1"/>"#,
            r#"<w:r><w:t>me}</w:t></w:r>"#,
            r#"<w:bookmarkEnd w:id="0"/>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert!(result.contains("Alice"), "placeholder with bookmark between runs not replaced: {}", result);
        assert!(!result.contains("{Na"), "leftover fragment: {}", result);
        assert!(result.contains("w:bookmarkStart"), "bookmark markers should be preserved: {}", result);
    }

    #[test]
    fn replace_with_comment_markers_between_runs() {
        let xml = concat!(
            r#"<w:r><w:t>{Na</w:t></w:r>"#,
            r#"<w:commentRangeStart w:id="1"/>"#,
            r#"<w:r><w:t>me}</w:t></w:r>"#,
            r#"<w:commentRangeEnd w:id="1"/>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert!(result.contains("Alice"), "placeholder with comment markers between runs not replaced: {}", result);
        assert!(!result.contains("{Na"), "leftover fragment: {}", result);
        assert!(result.contains("w:commentRangeStart"), "comment markers should be preserved: {}", result);
    }

    #[test]
    fn replace_with_formatting_props_between_runs() {
        let xml = concat!(
            r#"<w:r><w:rPr><w:b/></w:rPr><w:t>{Na</w:t></w:r>"#,
            r#"<w:r><w:rPr><w:i/></w:rPr><w:t>me}</w:t></w:r>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert!(result.contains("Alice"), "placeholder with rPr between runs not replaced: {}", result);
        assert!(!result.contains("{Na"), "leftover fragment: {}", result);
        assert!(result.contains("<w:rPr><w:b/></w:rPr>"), "formatting should be preserved: {}", result);
        assert!(result.contains("<w:rPr><w:i/></w:rPr>"), "formatting should be preserved: {}", result);
    }

    #[test]
    fn replace_with_empty_value() {
        let xml = r#"<w:p><w:r><w:t>Hello {Name}!</w:t></w:r></w:p>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "")]);
        assert_eq!(result, r#"<w:p><w:r><w:t>Hello !</w:t></w:r></w:p>"#);
    }

    #[test]
    fn replace_value_containing_curly_braces() {
        let xml = r#"<w:r><w:t>{Name}</w:t></w:r>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "{Alice}")]);
        assert_eq!(result, r#"<w:r><w:t>{Alice}</w:t></w:r>"#);

        let result = replace_placeholders_in_xml(xml, &[("{Name}", "a}b{c")]);
        assert_eq!(result, r#"<w:r><w:t>a}b{c</w:t></w:r>"#);
    }

    #[test]
    fn replace_with_multiline_value() {
        let xml = r#"<w:r><w:t>{Name}</w:t></w:r>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "line1\nline2\nline3")]);
        assert_eq!(result, r#"<w:r><w:t>line1
line2
line3</w:t></w:r>"#);
    }

    #[test]
    fn replace_same_placeholder_many_occurrences() {
        let xml = concat!(
            r#"<w:r><w:t>{x}</w:t></w:r>"#,
            r#"<w:r><w:t>{x}</w:t></w:r>"#,
            r#"<w:r><w:t>{x}</w:t></w:r>"#,
            r#"<w:r><w:t>{x}</w:t></w:r>"#,
            r#"<w:r><w:t>{x}</w:t></w:r>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{x}", "V")]);
        assert!(!result.contains("{x}"), "not all occurrences replaced: {}", result);
        assert_eq!(result.matches("V").count(), 5, "expected 5 replacements: {}", result);
    }

    #[test]
    fn drawingml_a_t_tags_are_replaced() {
        let xml = r#"<a:p><a:r><a:t>{placeholder}</a:t></a:r></a:p>"#;
        let result = replace_placeholders_in_xml(xml, &[("{placeholder}", "replaced")]);
        assert!(
            result.contains("replaced"),
            "DrawingML <a:t> tags should be replaced: {}",
            result
        );
        assert!(
            !result.contains("{placeholder}"),
            "DrawingML <a:t> placeholder should not remain: {}",
            result
        );
    }

    #[test]
    fn drawingml_a_t_split_across_runs() {
        let xml = r#"<a:r><a:t>{Na</a:t></a:r><a:r><a:t>me}</a:t></a:r>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert!(result.contains("Alice"), "split <a:t> placeholder not replaced: {}", result);
        assert!(!result.contains("{Na"), "leftover fragment: {}", result);
    }

    #[test]
    fn drawingml_a_t_escapes_xml() {
        let xml = r#"<a:t>{Name}</a:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice & Bob")]);
        assert_eq!(result, r#"<a:t>Alice &amp; Bob</a:t>"#);
    }

    #[test]
    fn wt_and_at_processed_independently() {
        let xml = r#"<w:r><w:t>{wt_val}</w:t></w:r><a:r><a:t>{at_val}</a:t></a:r>"#;
        let result = replace_placeholders_in_xml(
            xml,
            &[("{wt_val}", "Word"), ("{at_val}", "Drawing")],
        );
        assert!(result.contains("Word"), "w:t not replaced: {}", result);
        assert!(result.contains("Drawing"), "a:t not replaced: {}", result);
        assert!(!result.contains("{wt_val}"), "w:t placeholder remains: {}", result);
        assert!(!result.contains("{at_val}"), "a:t placeholder remains: {}", result);
    }

    #[test]
    fn math_m_t_tags_replaced() {
        let xml = r#"<m:r><m:t>{formula}</m:t></m:r>"#;
        let result = replace_placeholders_in_xml(xml, &[("{formula}", "x+1")]);
        assert_eq!(result, r#"<m:r><m:t>x+1</m:t></m:r>"#);
    }

    #[test]
    fn drawingml_a_t_with_attributes() {
        let xml = r#"<a:t xml:space="preserve">{placeholder}</a:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{placeholder}", "value")]);
        assert_eq!(result, r#"<a:t xml:space="preserve">value</a:t>"#);
    }

    // -- Tag boundary validation tests --
    // Ensures <w:t, <a:t, <m:t prefixes don't false-match longer tag names

    #[test]
    fn wt_prefix_does_not_match_w_tab() {
        let xml = r#"<w:r><w:tab/><w:t>{Name}</w:t></w:r>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert_eq!(result, r#"<w:r><w:tab/><w:t>Alice</w:t></w:r>"#);
    }

    #[test]
    fn wt_prefix_does_not_match_w_tbl() {
        let xml = r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>{Val}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Val}", "OK")]);
        assert!(result.contains("OK"), "placeholder not replaced: {}", result);
        assert!(!result.contains("{Val}"), "placeholder remains: {}", result);
    }

    #[test]
    fn at_prefix_does_not_match_a_tab() {
        let xml = r#"<a:p><a:r><a:tab/><a:t>{Name}</a:t></a:r></a:p>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert!(result.contains("<a:tab/>"), "a:tab should be untouched: {}", result);
        assert!(result.contains("Alice"), "placeholder not replaced: {}", result);
    }

    #[test]
    fn at_prefix_does_not_match_a_tbl_or_a_tc() {
        let xml = concat!(
            r#"<a:tbl><a:tr><a:tc><a:txBody>"#,
            r#"<a:p><a:r><a:t>{Cell}</a:t></a:r></a:p>"#,
            r#"</a:txBody></a:tc></a:tr></a:tbl>"#,
        );
        let result = replace_placeholders_in_xml(xml, &[("{Cell}", "Data")]);
        assert!(result.contains("Data"), "placeholder not replaced: {}", result);
        assert!(!result.contains("{Cell}"), "placeholder remains: {}", result);
    }

    #[test]
    fn self_closing_tags_are_skipped() {
        let xml = r#"<a:t/><a:t>{Name}</a:t>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Alice")]);
        assert!(result.contains("<a:t/>"), "self-closing tag should be untouched: {}", result);
        assert!(result.contains("Alice"), "placeholder not replaced: {}", result);
    }

    #[test]
    fn mt_prefix_does_not_match_longer_math_tags() {
        let xml = r#"<m:type>ignored</m:type><m:r><m:t>{X}</m:t></m:r>"#;
        let result = replace_placeholders_in_xml(xml, &[("{X}", "42")]);
        assert!(result.contains("ignored"), "m:type content should be untouched: {}", result);
        assert!(result.contains("42"), "placeholder not replaced: {}", result);
    }

    #[test]
    fn mixed_similar_tags_only_replaces_correct_ones() {
        let xml = concat!(
            r#"<w:tab/>"#,
            r#"<w:tbl><w:tr><w:tc></w:tc></w:tr></w:tbl>"#,
            r#"<w:r><w:t>{word}</w:t></w:r>"#,
            r#"<a:tab/>"#,
            r#"<a:tbl><a:tr><a:tc></a:tc></a:tr></a:tbl>"#,
            r#"<a:r><a:t>{draw}</a:t></a:r>"#,
            r#"<m:r><m:t>{math}</m:t></m:r>"#,
        );
        let result = replace_placeholders_in_xml(
            xml,
            &[("{word}", "W"), ("{draw}", "D"), ("{math}", "M")],
        );
        assert!(result.contains("<w:tab/>"), "w:tab modified");
        assert!(result.contains("<a:tab/>"), "a:tab modified");
        assert_eq!(result.matches("W").count(), 1);
        assert_eq!(result.matches("D").count(), 1);
        assert_eq!(result.matches("M").count(), 1);
        assert!(!result.contains("{word}"));
        assert!(!result.contains("{draw}"));
        assert!(!result.contains("{math}"));
    }

    #[test]
    fn prefix_at_end_of_string_does_not_panic() {
        let xml = "some text<a:t";
        let result = replace_placeholders_in_xml(xml, &[("{x}", "y")]);
        assert_eq!(result, xml);
    }

    #[test]
    fn w_t_with_space_preserve_attribute() {
        let xml = r#"<w:r><w:t xml:space="preserve"> {Name} </w:t></w:r>"#;
        let result = replace_placeholders_in_xml(xml, &[("{Name}", "Bob")]);
        assert!(result.contains("Bob"), "placeholder not replaced: {}", result);
    }

    fn create_test_zip(files: &[(&str, &[u8])]) -> Vec<u8> {
        let mut buf = Cursor::new(Vec::new());
        {
            let mut zip = zip::write::ZipWriter::new(&mut buf);
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            for &(name, content) in files {
                zip.start_file(name, options).unwrap();
                zip.write_all(content).unwrap();
            }
            zip.finish().unwrap();
        }
        buf.into_inner()
    }

    #[test]
    fn build_docx_replaces_in_footnotes_xml() {
        let footnotes_xml = concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?>"#,
            r#"<w:footnotes>"#,
            r#"<w:footnote w:id="1"><w:p><w:r><w:t>{Source}</w:t></w:r></w:p></w:footnote>"#,
            r#"</w:footnotes>"#,
        );
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?><w:document><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>"#;
        let template = create_test_zip(&[
            ("word/document.xml", doc_xml.as_bytes()),
            ("word/footnotes.xml", footnotes_xml.as_bytes()),
        ]);
        let result = __private::build_docx_bytes(&template, &[("{Source}", "Wikipedia")]).unwrap();
        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut xml = String::new();
        archive.by_name("word/footnotes.xml").unwrap().read_to_string(&mut xml).unwrap();
        assert!(xml.contains("Wikipedia"), "placeholder in footnotes.xml not replaced: {}", xml);
        assert!(!xml.contains("{Source}"), "placeholder still present: {}", xml);
    }

    #[test]
    fn build_docx_replaces_in_endnotes_xml() {
        let endnotes_xml = concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?>"#,
            r#"<w:endnotes>"#,
            r#"<w:endnote w:id="1"><w:p><w:r><w:t>{Citation}</w:t></w:r></w:p></w:endnote>"#,
            r#"</w:endnotes>"#,
        );
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?><w:document><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>"#;
        let template = create_test_zip(&[
            ("word/document.xml", doc_xml.as_bytes()),
            ("word/endnotes.xml", endnotes_xml.as_bytes()),
        ]);
        let result = __private::build_docx_bytes(&template, &[("{Citation}", "Doe 2024")]).unwrap();
        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut xml = String::new();
        archive.by_name("word/endnotes.xml").unwrap().read_to_string(&mut xml).unwrap();
        assert!(xml.contains("Doe 2024"), "placeholder in endnotes.xml not replaced: {}", xml);
        assert!(!xml.contains("{Citation}"), "placeholder still present: {}", xml);
    }

    #[test]
    fn build_docx_replaces_in_comments_xml() {
        let comments_xml = concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?>"#,
            r#"<w:comments>"#,
            r#"<w:comment w:id="0"><w:p><w:r><w:t>{Note}</w:t></w:r></w:p></w:comment>"#,
            r#"</w:comments>"#,
        );
        let doc_xml = r#"<?xml version="1.0" encoding="UTF-8"?><w:document><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>"#;
        let template = create_test_zip(&[
            ("word/document.xml", doc_xml.as_bytes()),
            ("word/comments.xml", comments_xml.as_bytes()),
        ]);
        let result = __private::build_docx_bytes(&template, &[("{Note}", "Approved")]).unwrap();
        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut xml = String::new();
        archive.by_name("word/comments.xml").unwrap().read_to_string(&mut xml).unwrap();
        assert!(xml.contains("Approved"), "placeholder in comments.xml not replaced: {}", xml);
        assert!(!xml.contains("{Note}"), "placeholder still present: {}", xml);
    }

    #[test]
    fn build_docx_replaces_across_multiple_xml_files() {
        let doc_xml = r#"<?xml version="1.0"?><w:document><w:body><w:p><w:r><w:t>{Body}</w:t></w:r></w:p></w:body></w:document>"#;
        let header_xml = r#"<?xml version="1.0"?><w:hdr><w:p><w:r><w:t>{Header}</w:t></w:r></w:p></w:hdr>"#;
        let footer_xml = r#"<?xml version="1.0"?><w:ftr><w:p><w:r><w:t>{Footer}</w:t></w:r></w:p></w:ftr>"#;
        let footnotes_xml = r#"<?xml version="1.0"?><w:footnotes><w:footnote w:id="1"><w:p><w:r><w:t>{FNote}</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
        let template = create_test_zip(&[
            ("word/document.xml", doc_xml.as_bytes()),
            ("word/header1.xml", header_xml.as_bytes()),
            ("word/footer1.xml", footer_xml.as_bytes()),
            ("word/footnotes.xml", footnotes_xml.as_bytes()),
        ]);
        let result = __private::build_docx_bytes(
            &template,
            &[("{Body}", "Main"), ("{Header}", "Top"), ("{Footer}", "Bottom"), ("{FNote}", "Ref1")],
        ).unwrap();
        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        for (file, expected, placeholder) in [
            ("word/document.xml", "Main", "{Body}"),
            ("word/header1.xml", "Top", "{Header}"),
            ("word/footer1.xml", "Bottom", "{Footer}"),
            ("word/footnotes.xml", "Ref1", "{FNote}"),
        ] {
            let mut xml = String::new();
            archive.by_name(file).unwrap().read_to_string(&mut xml).unwrap();
            assert!(xml.contains(expected), "{} not replaced in {}: {}", placeholder, file, xml);
            assert!(!xml.contains(placeholder), "{} still present in {}: {}", placeholder, file, xml);
        }
    }

    #[test]
    fn build_docx_preserves_non_xml_files() {
        let doc_xml = r#"<w:document><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>"#;
        let image_bytes: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0xFF, 0xFE];
        let template = create_test_zip(&[
            ("word/document.xml", doc_xml.as_bytes()),
            ("word/media/image1.png", image_bytes),
        ]);
        let result = __private::build_docx_bytes(&template, &[]).unwrap();
        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut output_image = Vec::new();
        archive.by_name("word/media/image1.png").unwrap().read_to_end(&mut output_image).unwrap();
        assert_eq!(output_image, image_bytes, "binary content should be preserved unchanged");
    }

    #[test]
    fn build_docx_does_not_replace_in_non_xml() {
        let doc_xml = r#"<w:document><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:document>"#;
        let bin_content = b"some binary with {Name} placeholder text";
        let template = create_test_zip(&[
            ("word/document.xml", doc_xml.as_bytes()),
            ("word/embeddings/data.bin", bin_content),
        ]);
        let result = __private::build_docx_bytes(&template, &[("{Name}", "Alice")]).unwrap();
        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut output_bin = Vec::new();
        archive.by_name("word/embeddings/data.bin").unwrap().read_to_end(&mut output_bin).unwrap();
        assert_eq!(output_bin, bin_content.as_slice(), ".bin file should not have replacements applied");
    }

    #[test]
    fn build_docx_replaces_in_drawingml_xml() {
        let diagram_xml = concat!(
            r#"<?xml version="1.0" encoding="UTF-8"?>"#,
            r#"<dgm:dataModel>"#,
            r#"<dgm:ptLst><dgm:pt><dgm:t><a:bodyPr/><a:p><a:r><a:t>{shape_text}</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst>"#,
            r#"</dgm:dataModel>"#,
        );
        let doc_xml = r#"<?xml version="1.0"?><w:document><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>"#;
        let template = create_test_zip(&[
            ("word/document.xml", doc_xml.as_bytes()),
            ("word/diagrams/data1.xml", diagram_xml.as_bytes()),
        ]);
        let result = __private::build_docx_bytes(&template, &[("{shape_text}", "Replaced!")]).unwrap();
        let cursor = Cursor::new(&result);
        let mut archive = zip::ZipArchive::new(cursor).unwrap();
        let mut xml = String::new();
        archive.by_name("word/diagrams/data1.xml").unwrap().read_to_string(&mut xml).unwrap();
        assert!(xml.contains("Replaced!"), "placeholder in DrawingML data1.xml not replaced: {}", xml);
        assert!(!xml.contains("{shape_text}"), "placeholder still present: {}", xml);
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
